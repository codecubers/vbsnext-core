Class DictUtil
    
    ' Author: Dale Fugier <https://developer.rhino3d.com/authors/dale_fugier>
    ' Source: https://developer.rhino3d.com/guides/rhinoscript/vbscript-dictionaries/
    ' Description:
    '   Sorts a dictionary by either key or item
    ' Parameters:
    '   objDict - the dictionary to sort
    '   intSort - the field to sort (1=key, 2=item)
    ' Returns:
    '   A dictionary sorted by intSort
    ''' <summary>Sorts a dictionary by either key or item</summary>
	''' <param name="objDict">The dictionary to sort</param>
	''' <param name="intSort">The field to sort (1=key, 2=item)</param>
    Function SortDictionary(objDict, intSort)

        ' declare constants
        Const dictKey  = 1
        Const dictItem = 2

        ' declare our variables
        Dim strDict()
        Dim objKey
        Dim strKey,strItem
        Dim X,Y,Z

        ' get the dictionary count
        Z = objDict.Count

        ' we need more than one item to warrant sorting
        If Z > 1 Then
            ' create an array to store dictionary information
            ReDim strDict(Z,2)
            X = 0
            ' populate the string array
            For Each objKey In objDict
                strDict(X,dictKey)  = CStr(objKey)
                strDict(X,dictItem) = CStr(objDict(objKey))
                X = X + 1
            Next

            ' perform a a shell sort of the string array
            For X = 0 To (Z - 2)
            For Y = X To (Z - 1)
                If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
                    strKey  = strDict(X,dictKey)
                    strItem = strDict(X,dictItem)
                    strDict(X,dictKey)  = strDict(Y,dictKey)
                    strDict(X,dictItem) = strDict(Y,dictItem)
                    strDict(Y,dictKey)  = strKey
                    strDict(Y,dictItem) = strItem
                End If
            Next
            Next

            ' erase the contents of the dictionary object
            objDict.RemoveAll

            ' repopulate the dictionary with the sorted information
            For X = 0 To (Z - 1)
            objDict.Add strDict(X,dictKey), strDict(X,dictItem)
            Next

        End If
    End Function
End Class