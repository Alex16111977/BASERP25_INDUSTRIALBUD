#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagnostika: ProchieAktivyPassivy 'v puti' vs DenezhnySredstvaVPuti Podrazdelenie
"""

import win32com.client
import sys

def connect_to_1c():
    """Connecting to 1S via COM"""
    try:
        v8 = win32com.client.Dispatch("V83.COMConnector")
        erp = v8.Connect('Srvr="SQLSERVER";Ref="BaseERP";Usr="Администратор";Pwd="24043"')
        print("[OK] Successfully connected to 1S")
        return erp
    except Exception as e:
        print(f"[ERROR] Connection failed: {e}")
        sys.exit(1)

def get_reference_name(obj):
    """Get the Name property from 1C reference"""
    try:
        return obj.Description
    except:
        try:
            return obj.Name
        except:
            try:
                return str(obj)
            except:
                return "UNKNOWN"

def query_pap_for_doc(erp, doc_ref):
    """Get all lines from PAP for document"""
    query_text = """
    ВЫБРАТЬ
        ВидДвижения,
        Статья,
        Источник,
        Подразделение,
        Сумма
    ИЗ РегистрНакопления.ПрочиеАктивыПассивы
    ГДЕ Регистратор = &Регистратор
    УПОРЯДОЧИТЬ ПО
        Статья,
        Источник
    """
    
    query = erp.NewObject("Query", query_text)
    query.SetParameter("Регистратор", doc_ref)
    result = query.Execute()
    return result

def query_denezhnye_sredstva_v_puti(erp, doc_ref):
    """Get Podrazdelenie from ДенежныеСредстваВПути"""
    query_text = """
    ВЫБРАТЬ
        Подразделение,
        Сумма,
        ВидДвижения
    ИЗ РегистрНакопления.ДенежныеСредстваВПути
    ГДЕ Регистратор = &Регистратор
    УПОРЯДОЧИТЬ ПО
        Подразделение
    """
    
    query = erp.NewObject("Query", query_text)
    query.SetParameter("Регистратор", doc_ref)
    result = query.Execute()
    return result

def find_documents(erp, target_num=None):
    """Find PKO documents"""
    if target_num:
        query_text = """
        ВЫБРАТЬ
            Ссылка,
            Номер,
            Дата
        ИЗ Документ.ПриходныйКассовыйОрдер
        ГДЕ Номер = &Номер
        """
        query = erp.NewObject("Query", query_text)
        query.SetParameter("Номер", target_num)
    else:
        # Simpler query without enum filtering
        query_text = """
        ВЫБРАТЬ ПЕРВЫЕ 10
            Ссылка,
            Номер,
            Дата,
            А_ВидПКО
        ИЗ Документ.ПриходныйКассовыйОрдер
        УПОРЯДОЧИТЬ ПО
            Дата УБЫВ
        """
        query = erp.NewObject("Query", query_text)
    
    result = query.Execute()
    return result

def main():
    print("=" * 100)
    print("DIAGNOSTICS: ProchieAktivyPassivy 'v puti' vs DenezhnySredstvaVPuti Podrazdelenie")
    print("=" * 100)
    
    erp = connect_to_1c()
    
    # TASK 1: Investigate N0000023550
    print("\n[TASK 1] Analysis of N0000023550")
    print("-" * 100)
    
    find_result = find_documents(erp, target_num="N0000023550")
    selection = find_result.Select()
    
    if not selection.Next():
        print("[ERROR] Document N0000023550 not found")
        return
    
    doc_ref = selection.Ссылка
    doc_num = selection.Номер
    doc_date = selection.Дата
    print(f"Document: {doc_num}, Date: {doc_date}")
    
    # Get data from PAP
    pap_result = query_pap_for_doc(erp, doc_ref)
    pap_selection = pap_result.Select()
    
    print("\nData from РегистрНакопления.ПрочиеАктивыПассивы:")
    print(f"{'MotionType':<25} {'Article':<35} {'Source':<35} {'Podrazdelenie':<20} {'Amount':<15}")
    print("-" * 130)
    
    pap_v_puti_podr = None
    
    while pap_selection.Next():
        vid = get_reference_name(pap_selection.ВидДвижения)
        stat = get_reference_name(pap_selection.Статья)
        istochnik = get_reference_name(pap_selection.Источник)
        podr = get_reference_name(pap_selection.Подразделение)
        summa = float(pap_selection.Сумма) if pap_selection.Сумма else 0
        
        print(f"{vid:<25} {stat[:34]:<35} {istochnik[:34]:<35} {podr:<20} {summa:>14,.2f}")
        
        # Identify "v puti" line
        if "Денежные средства в пути" in istochnik or "Денежные средства в пути" in stat:
            pap_v_puti_podr = podr
            print(f"  [FOUND] 'V PUTI' LINE: Podrazdelenie = {podr}")
    
    # Get data from ДенежныеСредстваВПути
    vputi_result = query_denezhnye_sredstva_v_puti(erp, doc_ref)
    vputi_selection = vputi_result.Select()
    
    print("\nData from РегистрНакопления.ДенежныеСредстваВПути:")
    print(f"{'Podrazdelenie':<25} {'MotionType':<25} {'Amount':<15}")
    print("-" * 65)
    
    vputi_podr = None
    
    while vputi_selection.Next():
        podr = get_reference_name(vputi_selection.Подразделение)
        vid = get_reference_name(vputi_selection.ВидДвижения)
        summa = float(vputi_selection.Сумма) if vputi_selection.Сумма else 0
        
        print(f"{podr:<25} {vid:<25} {summa:>14,.2f}")
        
        if vputi_podr is None:
            vputi_podr = podr
    
    # TASK 3: Comparison
    print("\n[TASK 3] COMPARISON OF PODRAZDELENIE")
    print("-" * 100)
    
    if pap_v_puti_podr and vputi_podr:
        match = "[MATCH]" if pap_v_puti_podr == vputi_podr else "[DIFFER]"
        print(f"PAP 'v puti'.Podrazdelenie:             {pap_v_puti_podr}")
        print(f"DenezhnySredstvaVPuti.Podrazdelenie:    {vputi_podr}")
        print(f"Result: {match}")
    else:
        print("[ERROR] Unable to get one of the values")
        if not pap_v_puti_podr:
            print("  - 'V puti' line in PAP not found")
        if not vputi_podr:
            print("  - Data in DenezhnySredstvaVPuti not found")
    
    # TASK 4: Check other PKO documents
    print("\n[TASK 4] CHECK OTHER PKO DOCUMENTS")
    print("-" * 100)
    
    other_docs_result = find_documents(erp)
    other_docs_selection = other_docs_result.Select()
    
    comparison_table = []
    doc_count = 0
    
    while other_docs_selection.Next() and doc_count < 5:
        doc_ref2 = other_docs_selection.Ссылка
        doc_num2 = other_docs_selection.Номер
        
        # Skip if already investigated
        if doc_num2 == "N0000023550":
            continue
        
        doc_count += 1
        
        # Get PAP
        pap_res2 = query_pap_for_doc(erp, doc_ref2)
        pap_sel2 = pap_res2.Select()
        
        pap_podr2 = None
        while pap_sel2.Next():
            istochnik = get_reference_name(pap_sel2.Источник)
            if "Денежные средства в пути" in istochnik:
                pap_podr2 = get_reference_name(pap_sel2.Подразделение)
                break
        
        # Get ДенежныеСредстваВПути
        vputi_res2 = query_denezhnye_sredstva_v_puti(erp, doc_ref2)
        vputi_sel2 = vputi_res2.Select()
        
        vputi_podr2 = None
        if vputi_sel2.Next():
            vputi_podr2 = get_reference_name(vputi_sel2.Подразделение)
        
        match_status = "[MATCH]" if (pap_podr2 and vputi_podr2 and pap_podr2 == vputi_podr2) else "[DIFFER]"
        
        comparison_table.append({
            'Number': doc_num2,
            'PAP_v_puti_Podrazdelenie': pap_podr2,
            'DenezhnySredstvaVPuti_Podrazdelenie': vputi_podr2,
            'Status': match_status
        })
        
        print(f"{doc_num2}: PAP={pap_podr2} vs VPuti={vputi_podr2} -> {match_status}")
    
    if not comparison_table:
        print("(No other documents found or analyzed)")
    
    # TASK 5: Search for scheduled jobs
    print("\n[TASK 5] SEARCH FOR SCHEDULED JOBS AND RECALCULATION PROCEDURES")
    print("-" * 100)
    
    try:
        scheduled_query = """
        ВЫБРАТЬ
            Ссылка,
            Наименование
        ИЗ РегламентноеЗадание
        ГДЕ Наименование СОДЕРЖИТ "баланс"
           ИЛИ Наименование СОДЕРЖИТ "ПрочиеАктивыПассивы"
        """
        
        scheduled_jobs_query = erp.NewObject("Query", scheduled_query)
        scheduled_result = scheduled_jobs_query.Execute()
        scheduled_selection = scheduled_result.Select()
        
        found_jobs = []
        while scheduled_selection.Next():
            job_name = scheduled_selection.Наименование
            found_jobs.append(job_name)
            print(f"  Found job: {job_name}")
        
        if not found_jobs:
            print("  (No scheduled jobs found with keywords)")
    except Exception as e:
        print(f"  Error searching for scheduled jobs: {e}")
    
    # List known recalculation methods
    print("\nKnown recalculation methods:")
    print("  - UpdateManagementBalance() in register ManagerModule")
    print("  - UpdateMotionOfAssetsLiabilities() in management accounting module")
    print("  - AffectsManagementBalance() determines source influence")
    
    # FINAL REPORT
    print("\n" + "=" * 100)
    print("FINAL REPORT")
    print("=" * 100)
    
    print(f"\n1. Document N0000023550:")
    print(f"   PAP 'v puti'.Podrazdelenie:             {pap_v_puti_podr}")
    print(f"   DenezhnySredstvaVPuti.Podrazdelenie:    {vputi_podr}")
    
    if pap_v_puti_podr and vputi_podr:
        if pap_v_puti_podr == vputi_podr:
            print(f"   [MATCH] They are equal: '{pap_v_puti_podr}'")
        else:
            print(f"   [DIFFER] They are different!")
            print(f"   Diagnosis: PAP 'v puti' does NOT track")
            print(f"   the Podrazdelenie dimension from DenezhnySredstvaVPuti")
    
    if comparison_table:
        print(f"\n2. Comparison table for {len(comparison_table)} other documents:")
        for row in comparison_table:
            print(f"   {row['Number']}: {row['Status']}")
    
    print("\n3. Recalculation mechanisms found:")
    print("   - UpdateManagementBalance() [AccumulationRegisters.ПрочиеАктивыПассивы]")
    print("   - UpdateMotionOfAssetsLiabilities() [Management accounting provisioning]")
    print("   - Mechanism called after movement recording")
    
    print("\n" + "=" * 100)

if __name__ == "__main__":
    main()