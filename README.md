Parse_Str
=========

Description:
Parse string is an Excel macro that parses string delimited by two different delimiter.


Syntax:
parse_str(string, left delimiter, right delimiter)


Example:
If cell A1 contains string, "cat dog <mouse> horse".  To parse text between delimiter '< >' from string in cell A1, type
  =parse_str(A1, "<", ">")

Cell A2 will have value "mouse"
