Attribute VB_Name = "demo_NoSQL"
Public Const demo_db_name As String = "test_nosqlite.nsql"
Public Const demo_db_collection As String = "contacts"

Private Sub test_nosqlite_TODO()

' s assurer du cast en str lors de chaque utilisation de .representation_dictionary.Item("_id")

' lors des updates, uniquement retourné les documents qui ont finalement effectivement reçus une modif et non pas ceux qui satisfont la query initiale

' implementation de l aggregation (lourd + complexe) -> passer par sql ?

' implementation des index

' update modifier $pull sur requete plutot que valeur

End Sub



Private Sub test_nosqlite_demo_queries()

'library VBA <-> JSON : passage sans perte d un dictionnaire Excel vers une notation plate JSON et vice-versa
    ' remarque importante : lorsque un array de text JSON est parse pour re-devenir un dictionnaire Excel, il devient un objet du type collection à la place d un array()
Dim oJSON As New jsonlib

'creation / chargement d'une DB
Dim oDBNoSQL As New cls_NoSQL_Database
    
    'si la base est nouvelle creation d un peu de contenu aleatoire
    If exist_file(ThisWorkbook.path & "\" & demo_db_name) = False Then
        MsgBox ("The system will now create a sample DB with 500 entries (contacts). This will take about 1min.")
        test_nosqlite_insert_random_data ThisWorkbook.path & "\" & demo_db_name, demo_db_collection
    End If

oDBNoSQL.setup_with_file ThisWorkbook.path & "\" & "test_nosqlite.nsql"



'creation / chargement d'une collection (equivalent d une table SQL)
Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use("contacts") 'choix de la collection ~ table



' ########### requetes d insertion de documents, function .insert des objets cls_NoSQL_Collection ###########
' parametre 1 [object/json] : document a inserer sous la forme d un dictionnaire ou de notation JSON
' parametre 2 [boolean] by default False : verifie en s assurant que la nouvelle entree est bien presente dans la base de donnees a l aide d un query de selection SQL

'ajout d un document

Dim dic_tmp_contact As New Scripting.Dictionary
    dic_tmp_contact.Add "name", "C"
    dic_tmp_contact.Add "surname", "g"
    dic_tmp_contact.Add "age", 30

    'insertion grace un document sous la forme d un dictionnaire Excel
    'oNoSQLCollection.insert dic_tmp_contact
    
    'insertion avec la notation JSON, les keys et valeur texte doivent etre entourees d apostrophe ou de DOUBLE guillemets
    'oNoSQLCollection.insert "{'name':'C',""surname"":'g', 'age':30}"



' ########### requetes de recherche, function .find des objets cls_NoSQL_Collection ###########

Dim tmp_oid As Variant 'key dico lors parcours resultat
Dim tmp_doc As cls_NoSQL_Document


Dim oResultSimple1FieldWithDico As cls_NoSQL_QueryResult
    Dim query_json As New Scripting.Dictionary
    query_json.Add "name", "A"
    Set oResultSimple1FieldWithDico = oNoSQLCollection.find(query_json)

    'parcours resultat : boucle sur le dictionaire .documents d un objet du type cls_NoSQL_QueryResult
'    For Each tmp_oid In oResultSimple1FieldWithDico.documents.keys
'        Set tmp_doc = oResultSimple1FieldWithDico.documents.Item(tmp_oid)
'
'        Debug.Print tmp_doc.representation_json
'    Next
    

' requete de recherche d egalite sur un seul field : {'key' : value}
Dim oResultSimple1FieldWithJSON As cls_NoSQL_QueryResult
    Set oResultSimple1FieldWithJSON = oNoSQLCollection.find("{'name':'A'}") 'equivalent a la requete precedente sans le passage par un dictionnaire


' trie d un resultat obtenu d une recherche ASCENDING : [ ['field', 1] ] ou DESCENDING [ ['field', -1] ]
    oResultSimple1FieldWithJSON.sort ("[['age', 1]]")
    
    'parcours resultat
'    For i = 0 To UBound(oResultSimple1FieldWithJSON.orders_keys, 1) 'boucle sur le vecteur .orders_keys <----
'        Set tmp_doc = oResultSimple1FieldWithJSON.documents.Item(oResultSimple1FieldWithJSON.orders_keys(i))
'
'        Debug.Print tmp_doc.representation_json
'    Next





' requete de recherche d egalite sur un field d un sous objet (objet d objet ou liste d objets contenue dans un array [...] : {'level_1.level_2.level_3 ... .level_n' : value}
Dim oResultSimpleSubFieldJSON As cls_NoSQL_QueryResult
    Set oResultSimpleSubFieldJSON = oNoSQLCollection.find("{'fb.username' : 'L'}")
    

'requete de recherche AND, egalite sur plusieurs champs en meme temps : {'field_1': value_1, 'field_2' : value_2, ...}
Dim oResultANDMultipleFieldsJSON_1 As cls_NoSQL_QueryResult, oResultANDMultipleFieldsJSON_2 As cls_NoSQL_QueryResult
    Set oResultANDMultipleFieldsJSON_1 = oNoSQLCollection.find("{'name':'A', 'surname':'s'}")
    
    'variante avec query operator $and: {'$and' : [ {'field_1' : value_1}, {'field_2' : value_2}, ...]}
    Set oResultANDMultipleFieldsJSON_2 = oNoSQLCollection.find("{'$and' : [{'name':'A'}, {'surname':'s'}]}")
    

' requete de recherche OR, egalite sur au moins un des champs : {'$or' : [{'field_1' : value_1}, {'field_2' : value_2}, ...]}
Dim oResultORMultipleFieldsJSON As cls_NoSQL_QueryResult
    Set oResultORMultipleFieldsJSON = oNoSQLCollection.find("{'$or' : [{'name':'A'}, {'surname':'s'}]}")


' requete de recherche XOR, OR exclusif, egalite sur au moins une des conditions mais pas TOUTES en même temps, ex fromage ou dessert mais pas les 2 : {'$nor' : [{'field_1' : value_1}, {'field_2' : value_2}, ...]}
Dim oResultXORMultipleFieldsJSON As cls_NoSQL_QueryResult
    Set oResultXORMultipleFieldsJSON = oNoSQLCollection.find("{'$nor' : [{'name':'A'}, {'surname':'s'}]}")


' requete de recherche plus grand > : {'field' : {'$gt' : value} }
Dim oResultGreaterThanJSON As cls_NoSQL_QueryResult
    Set oResultGreaterThanJSON = oNoSQLCollection.find("{'age' : {'$gt' : 50} }")

' requete de recherche plus grand > : {'field' : {'$lt' : value} }
Dim oResultLessThanJSON As cls_NoSQL_QueryResult
    Set oResultLessThanJSON = oNoSQLCollection.find("{'age' : {'$lt' : 10} }")


' chainage des plusieurs conditions de recherche sur un meme champ : {'field' : {condition_1, condition_2, ... condition_n}
Dim oResultChainConditionsSameFieldJSON As cls_NoSQL_QueryResult
    'Set oResultChainConditionsSameFieldJSON = oNoSQLCollection.find("{'age' : {'$gt' : 5, '$lt' : 10} }")
    Set oResultChainConditionsSameFieldJSON = oNoSQLCollection.find("{'tels.type.mobile_details.screen_size' : {'$gt' : 5, '$lt' : 7} }")


' requete de recherche de non egalite : {'field' : {'$ne' : value} }
Dim oResultNotEqualJSON As cls_NoSQL_QueryResult
    Set oResultNotEqualJSON = oNoSQLCollection.find("{'tels.type.mobile_details.screen_size' : {'$ne' : 7} }")

' requete de recherche d egalite sur au moins une valeur d un array : {'field' : {'$in' : [value_1, value_2, ..., value_n] } }
Dim oResultInJSON As cls_NoSQL_QueryResult
    Set oResultInJSON = oNoSQLCollection.find("{'age' : {'$in' : [48, 49, 50]} }")


' requete de recherche qui s assure qu une valeur n appartient pas a un array : {'field' : {'$nin' : [value_1, value_2, ..., value_n] } }
Dim oResultNinJSON As cls_NoSQL_QueryResult
    Set oResultNinJSON = oNoSQLCollection.find("{'age' : {'$nin' : [48, 49, 50]} }")


' requete de recherche qui s assure que une cle / champ est bien defini dans un objet
Dim oResultExistsJSON As cls_NoSQL_QueryResult
    Set oResultExistsJSON = oNoSQLCollection.find("{'tels.type.mobile_details.screen_size' : {'$exists' : true}}")


' requete de recherche sur une cle qui pointe sur un array : {'field_on_array' : {'$in' : [value_1, value_2, ..., value_n] } }
Dim oResultInArrayJSON As cls_NoSQL_QueryResult
    Set oResultInArrayJSON = oNoSQLCollection.find("{'tels.type.mobile_details.tel_available_color_array_no_object': {'$in' : ['yellow']}}")


' requete de recherche a l aide d une expression reguliere : {'field' : {'$regex' : 'pattern'}}
Dim oResultRegexJSON As cls_NoSQL_QueryResult
    Set oResultRegexJSON = oNoSQLCollection.find("{'fb.email': {'$regex' : 'vr'}}")



' ########### requetes de mise a jour, function .update des objets cls_NoSQL_Collection ###########
' parametre 1 [objet/json] : query de selection des documents grace a un dictionnaire ou de la notation JSON
' parametre 2 [objet/json] : query de mise a jour ou nouveau document qui ecrasera l ancien
' parametre 3 [boolean], by default False : si False, ne remplace que le premier document trouve, si True remplace tous les documents satisfaisants la requete de selection

' return : un objet cls_NOSQL_QueryResult contenant un dictionnaire d objets du type cls_NoSQL_Document ayant recus une modification

' remplace un document par un autre totalement nouveau : Set <objet_QueryResult> = <object_collection>.update( <query_de_selection/recherche> , <new_document_or_JSON>, True)
Dim oResultUpdateJSON As cls_NoSQL_QueryResult
    'Set oResultUpdateJSON = oNoSQLCollection.update("{'name':'W', 'surname':'a', age:45}", "{'name':'WW', 'surname':'b', age:45}")
    ' si la requete de recherche ne produit aucun resultat, une nouvelle entree est creee dans la base de donnees




' les UPDATE OPERATORS $ :

' remplacer la valeur de quelques champs : {'$set' : {'field_1' : <new_value_1>, 'field_2' : <new_value_2>, ...} }
Dim oResultUpdateSetJSON As cls_NoSQL_QueryResult
    Set oResultUpdateSetJSON = oNoSQLCollection.update("{'name':'A'}", "{'$set' : {'age' : -1, 'tels' : []}}", True)
    'si une cle est manquante dans le document à modifier, elle sera automatiquement creee

' supprimer une cle / champ : {'$unset' : {'field_to_delete_1' : '', 'field_to_delete_2' : '', ...} }
Dim oResultUpdateUnsetJSON As cls_NoSQL_QueryResult
    Set oResultUpdateUnsetJSON = oNoSQLCollection.update("{'name':'B'}", "{'$unset' : {'age' : ''}}", True)


' renommage d une cle : {'$rename' : {'field_old_name_1' : 'field_new_name_1', 'field_old_name_2' : 'field_new_name_2', ...} }
Dim oResultUpdateRenameJSON As cls_NoSQL_QueryResult
    Set oResultUpdateRenameJSON = oNoSQLCollection.update("{'name':'C'}", "{'$rename' : {'fb.username' : 'fb.login'}}", True)


' augmenter une valeur numerique uniquement si elle est plus petite qu un certain plafond : {'$max' : {'field' : <plafond> } }
Dim oResultUpdateMaxJSON As cls_NoSQL_QueryResult
    Set oResultUpdateMaxJSON = oNoSQLCollection.update("{'name':'D'}", "{'$max' : {'internal_counter' : 10}}", True) 'change la valeur de la cle "internal_counter" a la valeur du plafond si la valeur actuelle est inferieure a ce meme plafond


' incrementer / decrementer un compteur : {'$inc' : {'field_1' : <qty_a_ajouter_ou_a_retirer>} }
Dim oResultUpdateIncJSON As cls_NoSQL_QueryResult
    Set oResultUpdateIncJSON = oNoSQLCollection.update("{'name':'E'}", "{'$inc' : {'fb.total_hours_on_the_network' : 5}}", True)


' @@@@@ operation sur les arrays / collections @@@@@
' rajouter un scalaire / objet a un array existant : {'$push': {'field_for_array': <new_element>}}
Dim oResultUpdatePushJSON As cls_NoSQL_QueryResult
    Set oResultUpdatePushJSON = oNoSQLCollection.update("{'name':'F'}", "{'$push': {'array_random_int': -1111}}", True)

' retirer un element d un array grace a son index (ex : premier element index=-1, dernier element index=1) : {'$pop' : {'field_for_array': index}}
Dim oResultUpdatePopJSON As cls_NoSQL_QueryResult
    Set oResultUpdatePopJSON = oNoSQLCollection.update("{'name':'G'}", "{'$pop': {'array_random_int': 1}}", True)
    ' contrairement a mongodb, la notion d index a ete generalisee dans cette implementation : pour supprimer l avant-dernier element utiliser l index 2

' retirer un ou des elements d un array satisfaisants une valeur (a implementer : sur une requete et non une valeur) : {'$pull' : {'field_for_array' : <value>}}
Dim oResultUpdatePullJSON As cls_NoSQL_QueryResult
    Set oResultUpdatePullJSON = oNoSQLCollection.update("{'name':'H'}", "{'$pull': {'fb.last_connection': '31.08.2014'}}", True)




' ########### requetes de suppression, function .remove des objets cls_NoSQL_Collection ###########
' parametre 1 dictionnaire / JSON : requete de selection

' return : objet cls_NoSQL_QueryResult contenant un dictionnaire d objet cls_NoSQL_Document des documents supprimes de la base de donnees

Dim oResultRemove As cls_NoSQL_QueryResult
    Set oResultRemove = oNoSQLCollection.remove("{'name':'W'}")





' ########### requetes d aggregation, function .aggregate des objets cls_NoSQL_Collection ###########

'implementation ULTRA LIGHT et catastrophique, il n est pour l instant pas possible de construire a la volee des nouveaux champs a partir de combinaisons exec total_gross : qty_vendu * prix_vente
' seul $avg, $sum, $count sont applicable sur un seul champ a la fois
' pas de requete d aggreation possible sur des sous-objets contenu dans des ARRAYS


' format de la requete sur un model mongodb
' [{'$match' : {<selection_query>}}, {'$group' : { '_id' : {'header_group_by_field_1' : '$field_1', ...}, aggregate_header_1 : {'$accumulator': '$field_1'},... } }]

' la composante $match est facultative. Si elle n est pas precisee tous les documents seront utilises

' $group :
    ' le sous-objets "_id" correspond aux champs GROUP BY en SQL. Chaque champ est compose d une cle entete libre couplee au nom effective du champ dans la collection precede d un symbole dollar $
    ' calcul d aggregation : entete_libre : {'$accumulator' : '$champ'} ex  age_moyen : {'$avg' : '$age'}
        ' $accumulator possible : '$avg', '$count', '$sum', '$max', '$min'
            ' remarque importante : les aggregations ne sont possibles que sur des champs NUMERIQUES

Dim oResultAggregateJSON As cls_NoSQL_QueryResult
    ' agreger par nom les documents avec age > -50 en calculant pour chaque combinaison de nom possible la somme des ages et la moyenne des heures passees sur facebook
    Set oResultAggregateJSON = oNoSQLCollection.aggregate("[{'$match' : {'age': {'$gt': -50}}}, {'$group' : {'_id' : {'group_by_name' : '$name', 'group_by_surname' : '$surname'}, count_entries : {'$count' : '$age'}, total_age : {'$sum' : '$age'}, avg_age : {'$avg' : '$age'}, max_age : {'$max' : '$age'}, min_age : {'$min' : '$age'}, avg_hour_fb : {'$avg' : '$fb.total_hours_on_the_network'}}}]")
    
    'parcours du resultat
    For Each tmp_oid In oResultAggregateJSON.documents.keys
        Set tmp_doc = oResultAggregateJSON.documents.Item(tmp_oid)

        Debug.Print tmp_doc.representation_json
    Next


End Sub










Private Sub test_nosqlite_insert_random_data(ByVal db_path As String, ByVal collection As String)

Dim oJSON As New jsonlib

Dim oDBNoSQL As New cls_NoSQL_Database
oDBNoSQL.setup_with_file db_path


Dim oNoSQLCollection As New cls_NoSQL_Collection
Set oNoSQLCollection = oDBNoSQL.use(collection)


Dim tmp_dic As New Scripting.Dictionary, sub_dic As New Scripting.Dictionary

Dim vec_name() As Variant, vec_surname() As Variant, vec_tel_format() As Variant, vec_tel_brand() As Variant
    vec_name = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N")
    vec_surname = Array("n", "o", "p", "q", "r", "s", "t", "u", "v", "w")
    vec_tel_format = Array("work", "private", "mobile")
    vec_tel_brand = Array("lg", "samsung", "aapl", "nokia")
    vec_tel_color = Array("black", "gray", "gold", "red", "yellow", "white")
    
For k = 1 To 500
    Randomize
    
    Set tmp_dic = New Scripting.Dictionary
            Randomize
        tmp_dic.Add "name", vec_name(CInt(Rnd() * UBound(vec_name, 1)))
            Randomize
        tmp_dic.Add "surname", vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
        
        
        
            Randomize
        tmp_nbre_array_random_int = CInt(10 * Rnd()) + 1
        
        Dim vec_array_random_int() As Variant
        
        For i = 0 To tmp_nbre_array_random_int
            Randomize
            
            ReDim Preserve vec_array_random_int(i)
            vec_array_random_int(i) = CInt(100 * Rnd())
        Next i
        
        tmp_dic.Add "array_random_int", vec_array_random_int
        
        For i = 0 To UBound(vec_array_random_int, 1)
            vec_array_random_int(i) = vec_array_random_int(i) - CInt(100 * Rnd())
        Next i
        
        
            Randomize
        tmp_nbre_array_random_str = CInt(7 * Rnd()) + 1
        
        Dim vec_array_random_str() As Variant
        
        For i = 0 To tmp_nbre_array_random_str
            Randomize
            
            rnd_nbre_char = CInt(4 * Rnd()) + 1
            
            Dim tmp_str_for_array As String
            
            tmp_str_for_array = ""
            For j = 0 To rnd_nbre_char
                
                Randomize
                
                If j Mod 2 = 0 Then
                    tmp_str_for_array = tmp_str_for_array & vec_name(CInt(Rnd() * UBound(vec_name, 1)))
                Else
                    tmp_str_for_array = tmp_str_for_array & vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
                End If
            Next j
            
            ReDim Preserve vec_array_random_str(i)
            vec_array_random_str(i) = tmp_str_for_array
        Next i
        
        tmp_dic.Add "array_random_str", vec_array_random_str
        
        
        For i = 0 To UBound(vec_array_random_str, 1)
            Randomize
            
            rnd_nbre_char = CInt(4 * Rnd()) + 1
            
            For j = 0 To rnd_nbre_char
                
                Randomize
                
                If j Mod 2 = 0 Then
                    vec_array_random_str(i) = vec_array_random_str(i) & vec_name(CInt(Rnd() * UBound(vec_name, 1)))
                Else
                    vec_array_random_str(i) = vec_array_random_str(i) & vec_surname(CInt(Rnd() * UBound(vec_surname, 1)))
                End If
                
            Next j
            
        Next i
        
        
        
            
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
            
            dic_tmp_tel.Add "subarray_random_int", vec_array_random_int
            dic_tmp_tel.Add "subarray_random_str", vec_array_random_str
            
            
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
                            
                            Dim tmp_vec_color() As Variant, tmp_color As String, tmp_vec_color_not_as_an_object() As Variant
                            
                            n = 0
                            For j = 0 To tmp_nbre_color
                                
                                Randomize
                                tmp_color = vec_tel_color(CInt(Rnd() * UBound(vec_tel_color, 1)))
                                
                                If n = 0 Then
                                    Set dic_tel_subsubsubdic = New Scripting.Dictionary
                                        dic_tel_subsubsubdic.Add "color", tmp_color
                                    ReDim Preserve tmp_vec_color(n)
                                    Set tmp_vec_color(n) = dic_tel_subsubsubdic
                                    
                                    ReDim Preserve tmp_vec_color_not_as_an_object(n)
                                    tmp_vec_color_not_as_an_object(n) = tmp_color
                                    
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
                                                
                                                ReDim Preserve tmp_vec_color_not_as_an_object(n)
                                                tmp_vec_color_not_as_an_object(n) = tmp_color
                                                
                                                n = n + 1
                                            End If
                                        End If
                                        
                                    Next m
                                    
                                End If
                                
                            Next j
                            
                            'dic_tel_subsubdic.Add "tel_available_color_array", tmp_vec_color
                            dic_tel_subsubdic.Add "tel_available_color_array_no_object", tmp_vec_color_not_as_an_object
                            
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
            sub_dic.Add "total_hours_on_the_network", CInt(Rnd() * 1000)
            sub_dic.Add "nbre_connections", CInt(Rnd() * 100)
            
            Dim vec_connection_date() As Variant
            Randomize
            rnd_nbre_char = CInt(4 * Rnd()) + 1
            For i = 0 To rnd_nbre_char
                Randomize
                ReDim Preserve vec_connection_date(i)
                vec_connection_date(i) = Date - (CInt(100 * Rnd()) + 1)
            Next i
            
            sub_dic.Add "last_connection", vec_connection_date
            
            
        tmp_dic.Add "fb", sub_dic
        
        tmp_dic.Add "internal_counter", 0
    
    
    'Debug.Print oJSON.toString(tmp_dic)
    
    oNoSQLCollection.insert tmp_dic
    
Next k


End Sub


Private Sub test_vdo()

Dim oVDO As New cls_VDO
Set oVDO = oVDO.initWithFile("blub")
oVDO.query ("query")

''Dim oVDo As cls_VDO
'
'
'    Dim tmp_dic As New Scripting.Dictionary
'
'        tmp_dic.Add VDO_Field.vdo_field_name, "blub"
'        tmp_dic.Add VDO_Field.vdo_field_type, VDO_FieldType.vdo_ft_integer
'
'
'
'Dim oVDO As New cls_VDO
'    oVDO.createDB ("blub")
''Dim oVDOsqlite As New cls_VDO_sqlite

End Sub
