Attribute VB_Name = "bas_NoSQL"
'Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)

Private Sub test_nosqlite_TODO()

' s assurer du cast en str lors de chaque utilisation de .representation_dictionary.Item("_id")

End Sub

Private Sub test_nosqlite_insert()

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts")


Dim tmp_dic As New Scripting.Dictionary

Dim vec_name() As Variant, vec_surname() As Variant
    vec_name = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    vec_surname = Array("n", "o", "p", "q", "r", "s", "t", "u", "v", "w")
    
    
For k = 0 To 500
    Randomize
    
    Set tmp_dic = New Scripting.Dictionary
            Randomize
        tmp_dic.Add "name", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            Randomize
        tmp_dic.Add "surname", vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
            Randomize
        tmp_nbre_tel = CInt(7 * Rnd()) + 1
        Dim vec_tel() As Variant
        For i = 0 To tmp_nbre_tel
            
            ReDim Preserve vec_tel(i)
            
            tmp_tel = ""
            For j = 0 To 9
                Randomize
                tmp_tel = tmp_tel & CStr(CInt(Rnd() * 9))
                
            Next j
            
            vec_tel(i) = tmp_tel
            
        Next i
        
        tmp_dic.Add "tel", vec_tel
        
            Randomize
        tmp_dic.Add "age", CInt(Rnd() * 100)
    
    
    oNoSQLCollection.insert tmp_dic
    
    
Next k


End Sub


Private Sub test_nosqlite_insert_embedding()

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts_embedding")


Dim tmp_dic As New Scripting.Dictionary, sub_dic As New Scripting.Dictionary

Dim vec_name() As Variant, vec_surname() As Variant
    vec_name = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    vec_surname = Array("n", "o", "p", "q", "r", "s", "t", "u", "v", "w")
    
    
For k = 0 To 500
    Randomize
    
    Set tmp_dic = New Scripting.Dictionary
            Randomize
        tmp_dic.Add "name", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            Randomize
        tmp_dic.Add "surname", vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
            Randomize
        tmp_nbre_tel = CInt(7 * Rnd()) + 1
        Dim vec_tel() As Variant
        For i = 0 To tmp_nbre_tel
            
            ReDim Preserve vec_tel(i)
            
            tmp_tel = ""
            For j = 0 To 9
                Randomize
                tmp_tel = tmp_tel & CStr(CInt(Rnd() * 9))
                
            Next j
            
            vec_tel(i) = tmp_tel
            
        Next i
        
        tmp_dic.Add "tel", vec_tel
        
            Randomize
        tmp_dic.Add "age", CInt(Rnd() * 100)
            
            Set sub_dic = New Scripting.Dictionary
            sub_dic.Add "fid", CInt(Rnd() * 10000)
            sub_dic.Add "username", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            sub_dic.Add "email", vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
        
        tmp_dic.Add "fb", sub_dic
        
    oNoSQLCollection.insert tmp_dic
    
    
Next k


End Sub




Private Sub test_nosqlite_insert_embedding_array()

Dim oJSON As New jsonlib

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts_embedding_array")


Dim tmp_dic As New Scripting.Dictionary, sub_dic As New Scripting.Dictionary

Dim vec_name() As Variant, vec_surname() As Variant, vec_tel_format() As Variant, vec_tel_brand() As Variant
    vec_name = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    vec_surname = Array("n", "o", "p", "q", "r", "s", "t", "u", "v", "w")
    vec_tel_format = Array("work", "private", "mobile")
    vec_tel_brand = Array("lg", "samsung", "aapl", "nokia")
    vec_tel_color = Array("black", "gray", "gold", "red", "yellow", "white")
    
For k = 0 To 500
    Randomize
    
    Set tmp_dic = New Scripting.Dictionary
            Randomize
        tmp_dic.Add "name", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            Randomize
        tmp_dic.Add "surname", vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
            Randomize
        tmp_nbre_tel = CInt(5 * Rnd()) + 1
        Dim vec_tel() As Variant
        For i = 0 To tmp_nbre_tel
            
            Randomize
            
            Dim dic_tmp_tel As Scripting.Dictionary, dic_tel_subdic As Scripting.Dictionary, dic_tel_subsubdic As Scripting.Dictionary, dic_tel_subsubsubdic As Scripting.Dictionary
                Set dic_tmp_tel = New Scripting.Dictionary
            
            
            ReDim Preserve vec_tel(i)
            
            tmp_tel = ""
            For j = 0 To 4
                Randomize
                tmp_tel = tmp_tel & CStr(CInt(Rnd() * 9))
                
            Next j
            
            dic_tmp_tel.Add "num", tmp_tel
            
                Set dic_tel_subdic = New Scripting.Dictionary
                    Randomize
                    Dim tmp_format As String
                    tmp_format = vec_tel_format(CInt(Rnd() * UBound(vec_tel_format, 1)))
                    dic_tel_subdic.Add "format", tmp_format
                    
                    If tmp_format = "mobile" Then
                        Set dic_tel_subsubdic = New Scripting.Dictionary
                            
                            Randomize
                            dic_tel_subsubdic.Add "brand", vec_tel_brand(CInt(Rnd() * UBound(vec_tel_brand, 1)))
                            
                            Randomize
                            dic_tel_subsubdic.Add "screen_size", 4 + CInt(Rnd() * 5)
                            
                            
                            Dim tmp_nbre_color As Integer
                                Randomize
                                tmp_nbre_color = CInt(3 * Rnd()) + 1
                            
                            Dim tmp_vec_color() As Variant, tmp_color As String
                            
                            n = 0
                            For j = 0 To tmp_nbre_color
                                
                                Randomize
                                tmp_color = vec_tel_color(CInt(Rnd() * UBound(vec_tel_color, 1)))
                                
                                If n = 0 Then
                                    Set dic_tel_subsubsubdic = New Scripting.Dictionary
                                        dic_tel_subsubsubdic.Add "color", tmp_color
                                    ReDim Preserve tmp_vec_color(n)
                                    Set tmp_vec_color(n) = dic_tel_subsubsubdic
                                    n = n + 1
                                Else
                                    
                                    For m = UBound(tmp_vec_color, 1) To 0 Step -1
                                        If tmp_vec_color(m).Item("color") = tmp_color Then
                                            Exit For
                                        Else
                                            If m = 0 Then
                                                Set dic_tel_subsubsubdic = New Scripting.Dictionary
                                                    dic_tel_subsubsubdic.Add "color", tmp_color
                                                ReDim Preserve tmp_vec_color(n)
                                                Set tmp_vec_color(n) = dic_tel_subsubsubdic
                                                n = n + 1
                                            End If
                                        End If
                                        
                                    Next m
                                    
                                End If
                                
                            Next j
                            
                            dic_tel_subsubdic.Add "tel_available_color_array", tmp_vec_color
                            
                        dic_tel_subdic.Add "mobile_details", dic_tel_subsubdic
                        
                        
                        
                    End If
                    
                    
            dic_tmp_tel.Add "type", dic_tel_subdic
            
            
            
            
            
            
            
            Set vec_tel(i) = dic_tmp_tel
            
        Next i
        
        tmp_dic.Add "tels", vec_tel
        
            Randomize
        tmp_dic.Add "age", CInt(Rnd() * 100)
            
            Set sub_dic = New Scripting.Dictionary
            sub_dic.Add "fid", CInt(Rnd() * 10000)
            sub_dic.Add "username", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            sub_dic.Add "email", vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1))) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
        
        tmp_dic.Add "fb", sub_dic
    
    
    'Debug.Print oJSON.toString(tmp_dic)
    
    oNoSQLCollection.insert tmp_dic
    
    
Next k


End Sub



Private Sub test_nosqlite_query()



Dim oJSON As New jsonlib

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts")

Dim tmp_doc As cls_NoSQL_Document

Dim query_json As New Scripting.Dictionary, subquery_json As New Scripting.Dictionary
            query_json.Add "name", "E"
        Dim oResultSimple1field As cls_NoSQL_QueryResult, oResultSimpleMultipleField As cls_NoSQL_QueryResult, oResultNumFieldRange As cls_NoSQL_QueryResult, oResultNumFieldRangeFilteredFields As cls_NoSQL_QueryResult, _
            oResultNumFieldRangeOrdered As cls_NoSQL_QueryResult, oResultUpdateFullReplace As cls_NoSQL_QueryResult, oResultUpdateModifers As cls_NoSQL_QueryResult
            
        Set oResultSimple1field = oNoSQLCollection.find(oJSON.toString(query_json))
        
        For Each tmp_oid In oResultSimple1field.documents.keys
            Set tmp_doc = oResultSimple1field.documents.Item(tmp_oid)
            
            'Debug.Print oJSON.toString(tmp_doc.representation_dictionary.Item("_id"))
        Next
        
            oResultSimple1field.sort ("[[""age"",-1],[""name"", -1]]")
            
            For Each tmp_oid In oResultSimple1field.orders_keys
                Set tmp_doc = oResultSimple1field.documents.Item(tmp_oid)
                
                'Debug.Print tmp_doc.representation_json
            Next
        
        
        Set query_json = New Scripting.Dictionary
            query_json.Add "name", "E"
            query_json.Add "surname", "r"
        Set oResultSimpleMultipleField = oNoSQLCollection.find(oJSON.toString(query_json))
        
        For Each tmp_entry In oResultSimpleMultipleField.documents.keys
            Set tmp_doc = oResultSimpleMultipleField.documents.Item(tmp_entry)

            'Debug.Print tmp_doc.representation_json
        Next
        
        
        Set query_json = New Scripting.Dictionary
        Set subquery_json = New Scripting.Dictionary
            subquery_json.Add "$gt", 26
            subquery_json.Add "$lt", 50
            query_json.Add "age", subquery_json
            
            'Debug.Print oJSON.toString(query_json)
        Set oResultNumFieldRange = oNoSQLCollection.find(oJSON.toString(query_json))
        
        Set oResultNumFieldRangeFilteredFields = oNoSQLCollection.find(oJSON.toString(query_json), , , "{""name"":1, ""age"":1, ""blub"":1}")
        
        'controle filtre champs
        For Each tmp_entry In oResultNumFieldRangeFilteredFields.documents.keys
            Set tmp_doc = oResultNumFieldRangeFilteredFields.documents.Item(tmp_entry)

            'Debug.Print tmp_doc.representation_json
        Next
        
        
        Set o = oJSON.parse("[""age"":1]")
        
        
'        Set oResultNumFieldRangeOrdered = oNoSQLCollection.find(oJSON.toString(query_json), , "[""age"":1]")
'        vec_ordered_keys = oNoSQLCollection.last_find_results_ordered_keys
'
'        For i = 0 To UBound(vec_ordered_keys, 1)
'            Set tmp_doc = oResultNumFieldRangeOrdered.Item(vec_ordered_keys(i))
'            Debug.Print tmp_doc.representation_json
'        Next i

        Dim tmp_dic As New Scripting.Dictionary
        Set tmp_dic = New Scripting.Dictionary
        tmp_dic.Add "name", "AaA"
        tmp_dic.Add "surname", "bbb"
        tmp_dic.Add "tel", Array("12345", "78945", "987654321")
        tmp_dic.Add "age", 52

        'Set oResultUpdateFullReplace = oNoSQLCollection.update("{""name"":""A""}", tmp_dic)
        
        Set oResultSimple1field = oNoSQLCollection.find("{""name"":""AaA""}")
        
        'Set oResultUpdateModifers = oNoSQLCollection.update("{""name"":""B""}", "{""$set"": {""age"":-2, ""tel"":[""11"",""44""]}}", True)
        
        Set oResultSimple1field = oNoSQLCollection.find("{""name"":""B""}")
        
        For Each tmp_entry In oResultSimple1field.documents.keys
            Set tmp_doc = oResultSimple1field.documents.Item(tmp_entry)

            Debug.Print tmp_doc.representation_json
        Next
        
End Sub



Private Sub test_nosqlite_query_embedding()



Dim oJSON As New jsonlib

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts_embedding")

Dim tmp_doc As cls_NoSQL_Document

Dim query_json As New Scripting.Dictionary, subquery_json As New Scripting.Dictionary

Dim oResultSimple1field As cls_NoSQL_QueryResult, oResultSimpleMultipleField As cls_NoSQL_QueryResult, oResultNumFieldRange As cls_NoSQL_QueryResult, oResultNumFieldRangeFilteredFields As cls_NoSQL_QueryResult, _
    oResultNumFieldRangeOrdered As cls_NoSQL_QueryResult, oResultUpdateFullReplace As cls_NoSQL_QueryResult, oResultUpdateModifers As cls_NoSQL_QueryResult
            

Set oResultSimple1field = oNoSQLCollection.find("{'fb.username':'A'}")

Set oResultSimple1field = oNoSQLCollection.find("{'fb.fid': {'$lt' : 5000}}")

Set oResultUpdateFullReplace = oNoSQLCollection.remove("{'name':'B'}")

Set oResultSimpleMultipleField = oNoSQLCollection.find("{'name':'B'}")


End Sub




Private Sub test_nosqlite_query_embedding_array()



Dim oJSON As New jsonlib

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts_embedding_array")

Dim tmp_doc As cls_NoSQL_Document

Dim query_json As New Scripting.Dictionary, subquery_json As New Scripting.Dictionary

Dim oResultSimple1field As cls_NoSQL_QueryResult, oResultSimpleMultipleField As cls_NoSQL_QueryResult, oResultNumFieldRange As cls_NoSQL_QueryResult, oResultNumFieldRangeFilteredFields As cls_NoSQL_QueryResult, _
    oResultNumFieldRangeOrdered As cls_NoSQL_QueryResult, oResultUpdateFullReplace As cls_NoSQL_QueryResult, oResultUpdateModifers As cls_NoSQL_QueryResult
            

'Set oResultSimple1field = oNoSQLCollection.find("{'name':'A'}")
'Set oResultSimple1field = oNoSQLCollection.find("{'tels.type.format':'private'}")
'Set oResultSimple1field = oNoSQLCollection.find("{'tels.type.mobile_details.brand':'lg'}")

'Set oResultSimple1field = oNoSQLCollection.update("{'name' : 'H', 'tels.type.mobile_details.screen_size': {'$gt' : 7}, 'tels.type.mobile_details.brand': 'lg'}", "{'$set': {'age':-2,'tels.type.format' : 'mobile2'}}", True)

'Set oResultSimple1field = oNoSQLCollection.find("{'name' : 'H', 'tels.type.mobile_details.screen_size': {'$gt' : 7}, 'tels.type.mobile_details.brand': 'lg'}")

Set oResultSimple1field = oNoSQLCollection.find("{'tels.type.mobile_details.tel_available_color_array.color': 'yellow'}")

For Each tmp_oid In oResultSimple1field.documents.keys
    Set tmp_doc = oResultSimple1field.documents.Item(tmp_oid)
    
    Debug.Print tmp_doc.representation_json
Next

End Sub



Private Sub test_nosqlite_dic()

Dim oJSON As New jsonlib

Set blub = oJSON.parse("[[""age"",1],[""name"", -1]]") ' retourne une collection car l ordre compte dans un array !

Dim vec_name() As Variant, vec_surname() As Variant
    vec_name = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    vec_surname = Array("n", "o", "p", "q", "r", "s", "t", "u", "v", "w")

Dim tmp_dic As Scripting.Dictionary
Set tmp_dic = New Scripting.Dictionary
            Randomize
        tmp_dic.Add "name", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            Randomize
        tmp_dic.Add "surname", vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
            Randomize
        tmp_nbre_tel = CInt(7 * Rnd()) + 1
        Dim vec_tel() As Variant
        For i = 0 To tmp_nbre_tel
            
            ReDim Preserve vec_tel(i)
            
            tmp_tel = ""
            For j = 0 To 9
                Randomize
                tmp_tel = tmp_tel & CStr(CInt(Rnd() * 9))
                
            Next j
            
            vec_tel(i) = tmp_tel
            
        Next i
        
        tmp_dic.Add "tel", vec_tel
        
            Randomize
        tmp_dic.Add "age", CInt(Rnd() * 100)

Debug.Print oJSON.toString(vec_tel)

End Sub


Private Sub test_nosqlite_manip_simple_object()
    
    Dim oNSOId As New cls_NoSQL_ObjectId
    
    For i = 0 To 10
        'Debug.Print oNSOId.get_next()
    Next i
    
    
    Dim oJSON As New jsonlib
    
    
    Set o = oJSON.parse("{'name':'B','surname':'s','tel':['11','44']}")
    
    Dim oDBNoSQL As New cls_NoSQL_Database
    oDBNoSQL.setup_with_file "s:\blp\data\test_nosqlite.nsql"
    
    
    Dim tmp_dic As New Scripting.Dictionary
    Set tmp_dic = New Scripting.Dictionary
        tmp_dic.Add "name", "AAA"
        tmp_dic.Add "surname", "bbb"
        tmp_dic.Add "tel", Array("12345", "78945", "987654321")
        tmp_dic.Add "age", 52
    
    
    k = tmp_dic.keys(0)
    
    Dim oDocNoSQL As New cls_NoSQL_Document, oDocNoSQL2 As New cls_NoSQL_Document
        oDocNoSQL.load_data tmp_dic
        oDocNoSQL2.load_data oJSON.toString(tmp_dic)
    
    Dim oDic As Scripting.Dictionary
    Set oDic = oJSON.parse(CStr(oJSON.toString(tmp_dic)))
        Set oSubOject = oDic.Item("tel") 'array est devenu un objet collection
    
    table_structure = sqlite3_get_table_structure("s:\blp\data\test_nosqlite.nsql", "helper")
    
    
End Sub


