Option Explicit
Const ForReading = 1, forAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

const functionPattern = "((\s?'.*\r\n)*)\r\n\s*(public|private)?\s*()\s*(function)\s+()([^'\r]+).*\r\n([\s\S]+?)\r\n\s*end function"
const classPattern = "((\s?'.*\r\n)*)\r\n\s*(public|private)?\s*()\s*(class)\s+()([^'\r]+).*\r\n([\s\S]+?)\r\n\s*end class"
const subPattern = "((\s?'.*\r\n)*)\r\n\s*(public|private)?\s*()\s*(sub)\s+()([^'\r]+).*\r\n([\s\S]+?)\r\n\s*end sub"
const propertyPattern = "((\s?'.*\r\n)*)\r\n\s*(public|private)?\s*(default)?\s*(property)\s+(set|get|let)\s+([^'\r]+).*\r\n([\s\S]+?)\r\n\s*end property"


main 

Sub main

	dim scriptToXml 	: 	Set scriptToXml = new cls_ScriptToXml
	
	scriptToXml.RootNode = "module"
	scriptToXml.FileName = Wscript.Arguments.Item(1)
	Set scriptToXml.ScriptObject = ParseScript (WScript.Arguments.Item(0))
	scriptToXml.ScriptObject.Add "destination", scriptToXml.FileName
	scriptToXml.Execute
	
End Sub

Public Function ParseScript(fileName)

	dim cls
	dim codeDictionary
	Dim classExtractor
		
	Set codeDictionary = CreateObject("Scripting.Dictionary")
	
	codeDictionary.Add "innercode", ReturnFileAsText(fileName)
	codeDictionary.Add "source", fileName
		
	' First, extract the class
	Set classExtractor = new_ClassExtractor
	
	' Add the class objects to the codedictionary
	CodeDictionary.Add "classcollection", classExtractor.CodeToCollection(CodeDictionary.Item("innercode"))
	
	' Parse the code inside all classes to function, sub and property objects
	for each cls in CodeDictionary.item("classcollection")
		cls.Add "functioncollection", (new_FunctionExtractor).CodeToCollection(cls.item("innercode"))
		cls.Add "subcollection", (new_SubExtractor).CodeToCollection(cls.item("innercode"))
		cls.Add "propertycollection", (new_PropertyExtractor).CodeToCollection(cls.item("innercode"))
	next
	
	' Parse the code outside the classes to function, sub and property objects
	CodeDictionary.Add "functioncollection", (new_FunctionExtractor).CodeToCollection(classExtractor.OuterCode)
	CodeDictionary.Add "subcollection", (new_subExtractor).CodeToCollection(classExtractor.OuterCode)
	CodeDictionary.Add "propertycollection", (new_PropertyExtractor).CodeToCollection(classExtractor.OuterCode)			
	
	Set ParseScript = CodeDictionary	
	
end Function

Class cls_ScriptToXml

	private xmlDoc_
	Private Sub Class_Initialize
		RootNode = "root"
		Set xmlDoc_ = CreateObject("Microsoft.XMLDOM")
	End Sub

	private rootNode_
	Public Property Let RootNode(p)
		rootNode_ = p
	End Property
	Public Property Get RootNode
		RootNode = rootNode_
	End Property
	
	Private fileName_
	Public Property Let FileName(p)
		fileName_ = p
	End Property
	Public Property Get FileName
		FileName = fileName_
	End Property
	
	Private scriptObject_
	Public Property Set ScriptObject(o)
		Set scriptObject_ = o
	End Property
	Public Property Get ScriptObject
		Set ScriptObject = scriptObject_
	End Property

	Public Sub Execute
		objToXmlElems xmlDoc_, RootNode, ScriptObject
		xmlDoc_.Save FileName
	End Sub
	
	Private Sub objToXmlElems(xmlObj, nodename, t)
		dim elem, k, child
		
		Select Case TypeName(t)
			Case "ArrayList"
				set elem = xmlDoc_.CreateElement(nodename)
				for each k in t
					objToXmlElems elem, replace(nodename, "collection", ""), k
				next
				xmlObj.appendChild elem
			
			Case "Dictionary"
				Set elem = xmlDoc_.CreateElement(nodename)
				for each k in t.keys
					objToXmlElems elem, k, t.item(k)
				next
				xmlObj.AppendChild elem
			
			Case Else
				Set elem = xmlDoc_.CreateElement(nodename)
				elem.Text = t
				xmlObj.AppendChild(elem)
			
		end select
	End Sub


End Class



Public Function ReturnFileAsText(filename)

	Dim objFSO, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(filename, ForReading, False, TristateUseDefault)
	
	ReturnFileAsText = objFile.ReadAll

End Function

Public Function new_FunctionExtractor
	Set new_FunctionExtractor = new cls_CodeExtractor
	new_FunctionExtractor.Pattern = functionPattern
End Function

Public Function new_SubExtractor
	Set new_SubExtractor = new cls_CodeExtractor
	new_SubExtractor.Pattern = subPattern
End Function

Public Function new_PropertyExtractor
	Set new_PropertyExtractor = new cls_CodeExtractor
	new_PropertyExtractor.Pattern = propertyPattern
End Function

Public Function new_ClassExtractor
	Set new_ClassExtractor = new cls_CodeExtractor
	new_ClassExtractor.Pattern = classPattern
End Function

Private Function new_ExtractorPrototype
	Dim o : set o = new cls_CodeExtractor
	o.AddRegexmatchParser 0, getref("topcommentParser")
	' etc.
	Set new_ExtractorPrototype = o
End Function


Class cls_CodeExtractor

	private originalCode_
	Public Property Get OriginalCode
		OriginalCode = originalCode_
	End Property
	Private Property Let OriginalCode(p)
		originalCode_ = p
	End Property
	
	private outerCode_
	Public Property Get OuterCode
		OuterCode = outerCode_
	End Property
	Private Property Let OuterCode(p)
		OuterCode_ = p
	End Property
	
	private pattern_
	Public Property Let Pattern(p)
		pattern_ = P
	End Property
	Public Property Get Pattern
		Pattern = pattern_
	End Property
	
	Public Function CodeToCollection(code)
		
		dim i, re, matches, match
		dim objCollection, contentDictionary
		
		OriginalCode = code
		
		Set objCollection = CreateObject("System.Collections.ArrayList")
		
		Set re = new regexp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = Pattern
				
		Set matches = re.Execute(code)
			
		for each match in Matches
			
			set contentDictionary = CreateObject("Scripting.Dictionary")
			
			contentDictionary.Add "topcomment", match.submatches(0)	' comment
			contentDictionary.Add "scope", match.submatches(2)	' comment
			contentDictionary.Add "default", match.submatches(3)	' comment
			contentDictionary.Add "type", match.submatches(4)	' comment
			contentDictionary.Add "getsetlet", match.submatches(5)	' comment
			contentDictionary.Add "nameandparams", match.submatches(6)	' comment
			contentDictionary.Add "innercode", match.submatches(7)	' comment
			
			objCollection.Add contentDictionary
		next
		
		OuterCode = re.Replace(code, "")
		
		Set CodeToCollection = objCollection		
	End Function

End Class


		
		
