Attribute VB_Name = "dxregex"
' Module created to Assist in parsing Dr Checks XML reports exported from ProjNet.
Public Const module_project As String = "dxreview"
Public Const module_name As String = "dxregex"
Public Const module_version As String = "1.0"
Public Const module_author As String = "Ben Fisher"
Public Const module_date As Date = #3/16/2024#

' Set a reference to:
'    Microsoft VBScript Regular Expressions 5.5 (vbscript.dll) - used for regular expressions (RegExp)
'    Microsoft Scripting Runtime (scrrun.dll) - used for Scripting.Dictionary

'========================================================================================
'
'                                    Regex Functions
'
'========================================================================================

Private Function DictionaryToString(a_dictionary As Dictionary, Optional delineator As String = vbCrLf)
    ' Write all the keys of a dictionary to a single string, delineated by the supplied character or string
    Dim a_string As String
    For Each a_key In a_dictionary.Keys
        If a_string = Empty Then a_string = a_key Else a_string = a_string & delineator & a_key
    Next
    DictionaryToString = a_string
End Function

Private Function GetAllMatches(search_string As String, pattern_array As Variant, Optional isGlobal As Boolean = True, Optional isCaseIgnored As Boolean = False, Optional isMultiLine As Boolean = True)
    ' Return a flattened list of all matches to sheet number like text inside a range of cells

    Dim regex As New RegExp
    Dim match_collection As MatchCollection
    Dim a_match As Match
    Dim match_dictionary As Dictionary
    Dim match_string As String
    Dim delineator As String
    Dim phrase As Variant

    ' This delineator will be used at the end with DictionaryToString to create a return string
    delineator = vbCrLf

    ' Set global parameters for the regex search
    With regex
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
    End With

    ' Perform a regex search for each pattern, use a Dictionary, which only permits unique keys
    ' latter we'll flatten the dictionary keys into a single return string.
    Set match_dictionary = New Dictionary
    For Each phrase In pattern_array
        regex.Pattern = phrase
        Set match_collection = regex.Execute(search_string)
        For Each a_match In match_collection
            If Not match_dictionary.Exists(a_match.Value) Then
                match_dictionary.Add a_match.Value, a_match.Value
            End If
        Next
    Next
    
    ' Flatten dictionary of matches and return that string
    GetAllMatches = DictionaryToString(match_dictionary, delineator)
End Function

Public Function Find_SheetNumbers(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches for sheet number patterns in specified search_cells (Excel ranges)
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("\b([AC-Z]-{1}|[AC-Z]{2})(\d{3})(-?[A-Z]{1,2})?")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_SheetNumbers = Replace(GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns), "DD139", "")
End Function

Public Function Find_ReportSectionNumbers(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches of report section number patterns in specified search_cells (Excel ranges)
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("((Section|section|Sec|sec)\.?\s?)?(\w)?(\d{1,4}-)?\d+([\.-]{1}\d+)+")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_ReportSectionNumbers = GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns, isCaseIgnored:=True)
End Function

Public Function Find_TablesFiguresPages(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches of report section number patterns in specified search_cells (Excel ranges)
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("(Table|Figure|Fig)\.?\s(\d{1,4}|[A-Z])?-?(\d{0,3}\.?\d{1,4})?", _
                     "\b(Page|page|p.|pp.)\s?\d+")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_TablesFiguresPages = GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns)
End Function

Public Function Find_StandardReferences(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches for codes and standards patterns in specified search_cells (Excel ranges)
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("DD1391", _
                     "UFC\s?\d{1,2}-\d{3}-\d{1,3}", _
                     "\bFC\s?\d{1,2}-\d{3}-\d{1,3}[A-Z]{1,2}", _
                     "MIL-STD-\d{1,4}[A-Z]{0,2}", _
                     "(\d{4}\s)?(ACI|AISC|ASCE|ASTM|TMS|CMAA|AREMA|NFPA|FEMA|ISO|IBC|IFC|IPC|ANSI|ASME|AISI|IEEE|NAVSEA|JARPA|NHPA|ARPA|ASHRAE|EPA)(\s\d{1,5}(-\d{1,3})?)?(\sand\s\d{1,4})?", _
                     "(JIS|JSA|JASS|WJES|JES|AIJ|MLIT|JSCA|JSCE)", _
                     "\d{1,3}\sCFR\s(Part|PART|part)?\s\d{1,5}")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_StandardReferences = GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns)
End Function

Public Function Find_SpecSections(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches for UFGS spec section patterns in specified search_cells (Excel ranges)
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("(UFGS\s?)?(\d{2}\s\d{2}\s\d{2})+(\.\d{2}(\s\d{2})?)?")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_SpecSections = GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns)
End Function

Public Function Find_BuildingAndRoomNumbers(ParamArray search_cells()) As String
    ' Return delineated string of all regex matches for building or room number like patterns
    ' in specified search_cells (Excel ranges). This particular fuction has limitations as
    ' some of the common human-made errors look like sheet numbers
    Dim patterns As Variant
    Dim combined_cell_values As String
    Dim thing As Variant
    
    ' Load up an array of regex search patterns and assign to variable
    patterns = Array("[BP]\s?-?\d{3,4}", _
                     "(Bldg|Building|bldg|building)\.?\s?-?\d{1,4}", _
                     "(Room|ROOM|room|rm)\s?[A-Z]?\d{1,3}")
    
    ' Combine the text from all user provided search cells into a single string
    For Each thing In search_cells
        combined_cell_values = combined_cell_values & " " & thing
    Next
    Find_BuildingAndRoomNumbers = GetAllMatches(search_string:=combined_cell_values, pattern_array:=patterns)
End Function

