Attribute VB_Name = "Module1"
'Rich Leung | 11/14/2013
'Parse word delimited by delimiter sep1 and sep2

Function Parse_Str(str, sep1, sep2) As String
    Dim V1() As String
    Dim V2() As String
    'Dim temp_str() As String
    
    V1 = Split(str, sep1)
    V2 = Split(V1(1), sep2)
    Parse_Str = V2(0)
    
End Function
