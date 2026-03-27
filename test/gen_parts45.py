import base64

apo = chr(39)
emdash = chr(8212)

# Part 4: KonvertirovatTZ_BuhBud
p4_lines = [
    chr(13)+chr(10),
    chr(13)+chr(10),
    "// =====================================================================" + chr(13)+chr(10),
]

# Build part 4 header comment
p4_lines.append("// " + chr(1050)+chr(1054)+chr(1053)+chr(1042)+chr(1045)+chr(1056)+chr(1058)+chr(1040)+chr(1062)+chr(1030)+chr(1071) + " BuhBud: " + chr(1040)+"_"+chr(1048)+chr(1076)+chr(1050)+chr(1086)+chr(1076) + " -> " + chr(1057)+chr(1087)+chr(1088)+chr(1072)+chr(1074)+chr(1086)+chr(1095)+chr(1085)+chr(1080)+chr(1082)+chr(1057)+chr(1089)+chr(1099)+chr(1083)+chr(1082)+chr(1072)+"."+chr(1060)+chr(1080)+chr(1079)+chr(1080)+chr(1095)+chr(1077)+chr(1089)+chr(1082)+chr(1080)+chr(1077)+chr(1051)+chr(1080)+chr(1094)+chr(1072) + chr(13)+chr(10))
print("P4 header generated")
