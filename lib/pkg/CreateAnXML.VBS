Dim doc     'as MSXML2.DOMDocument
Dim nRoot   'as MSXML2.IXMLDOMElement
Dim nPri    'as MSXML2.IXMLDOMElement
Dim nSec    'as MSXML2.IXMLDOMElement

Set doc = createobject("MSXML2.DOMDocument")  'Remove Createobject if referencing xml dll in vba project
Set nRoot = doc.createElement("SubjectVBS2")
Set nPri = doc.createElement("Primary")
    
    Set nPriNode = doc.createAttribute("Parameter")
    nPriNode.NodeValue = "Release"
    Set nPriNode2 = doc.createAttribute("DefaultValue")
    nPriNode2.NodeValue = "June Release 2012"
    
Set nSec = doc.createElement("Secondary")
    nSec.Text = "SecondNodeText"

doc.appendChild nRoot
nRoot.appendChild nPri
nPri.appendChild nSec
    nPri.setAttributeNode nPriNode
    nPri.setAttributeNode nPriNode2

doc.Save "c:\users\pravy\desktop\MyFirstXML_VBS.xml"
Set doc = Nothing