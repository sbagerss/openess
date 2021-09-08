import xml.etree.ElementTree as et

##################################
class Argument:
    element = None

    def __init__(self, element):
        self.element = element
    
    def show(self):
        ...

    def getArg(self):
        ...

    def setArg(self,text):
        ...

##################################
class ArgumentText(Argument):

    def __init__(self, element):
        super().__init__(element)
        
    def getArg(self):
        return self.element.text

    def setArg(self,text):
        self.element.text = text

    def show(self):
        print(self.element.text)

##################################
class ArgumentAttrib(Argument):
    attribName = ""

    def __init__(self, element, attribName):
        super().__init__(element)
        self.attribName = attribName

    def getArg(self):
        return self.element.get(self.attribName)

    def setArg(self, text):
        self.element.set(self.attribName,text)

    def show(self):
        print(self.element.get(self.attribName))

##################################
class CodeXml:
    code = None
    
    ##################################
    def __init__(self, code):
        self.code = code

    ##################################
    def findArgsWithSpecChar(self, character):
        return CodeXml._findArgsWithSpecCharRec(self.code,character)

    ##################################
    def _findArgsWithSpecCharRec(parent, character):

        argPosList = []

        for child in parent:

            for atr in child.attrib:

                atrValue = child.get(atr)
                if atrValue != None:
                    if len(atrValue) > 0:
                        if atrValue[0] == character:
                            argPosList = argPosList + [ArgumentAttrib(child,atr)]

            if(len(child) > 0):
                argPosList = argPosList + CodeXml._findArgsWithSpecCharRec(child,character)
            else:
                if(child.text != None):
                    if(child.text[0] == character):
                        argPosList = argPosList + [ArgumentText(child)]
                    
        return argPosList

    ##################################
    def completeArgs(self, argList):
        
        argListPos = self.findArgsWithSpecChar('$')

        i = 0
        for arg in argListPos:
            s = arg.getArg()
            if s != None:
                if s[0] == '$':
                    argNum = int(s[1:])
                    arg.setArg(argList[i])
            i = i + 1
        return 0

    ##################################
    def _getParentForInsertCode(parent):

        for child in parent:
            if(child.tag == 'code'):
                return parent

        for child in parent:
            ret = CodeXml._getParentForInsertCode(child)
            if(ret != None):
                return ret

    ##################################
    def insertCode(self, codeRoot):
        parent = CodeXml._getParentForInsertCode(self.code)

        i = 0
        for child in parent:
            if child.tag == 'code':
                if i > 0:
                    parent.insert(i,codeRoot)
                    break
            i = i + 1

    ##################################
    def _deleteCodeTag(self):
        parent = CodeXml._getParentForInsertCode(self.code)

        if parent != None:
            for child in parent:
                if child.tag == 'code':
                    parent.remove(child)
                    break

    ##################################
    def _setIds(self):
        argListPos = self.findArgsWithSpecChar('#')
        i = 0
        for arg in argListPos:
            s = arg.getArg()
            if s != None:
                if s[0] == '#':
                    arg.setArg(str(i))
            i = i + 1      

    ##################################
    def _setXmlns(parent):
        foundInterface = False
        foundFlgNet = False

        for child in parent:

            for at in child.attrib:
                if at == "xmlnsIntefaceTag":
                    foundInterface = True
                    break
                if at == 'xmlnsFlgNetTag':
                    foundFlgNet = True
                    break

            if foundInterface:
                child.attrib.pop("xmlnsIntefaceTag")
                child.attrib["xmlns"] = "http://www.siemens.com/automation/Openness/SW/Interface/v4"

            if foundFlgNet:
                child.attrib.pop("xmlnsFlgNetTag")
                child.attrib["xmlns"] = "http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v4"

        for child in parent:
            CodeXml._setXmlns(child)

    ##################################
    def prepareCode(self):
        self._deleteCodeTag()
        self._setIds()
        CodeXml._setXmlns(self.code)


##################################
fcTree = et.parse('C:\\pythonProject\\openess\\templates\\fc.xml')
rootFc = fcTree.getroot()

# blockCallTree = et.parse('C:\\pythonProject\\openess\\templates\\blokFunkcyjny.xml')
# blockCallRoot = blockCallTree.getroot()

x = CodeXml(rootFc)
# y = CodeXml(blockCallRoot)
# argList = ["zmienna1", "Data_block_1", "a1", "superZmienna", "Data_block_2", "y1", "Data_block_2" ,"y2", "DI_blokFunkcyjny_A10", "adasd"]
# y.completeArgs(argList)

i = 100
while i > 0:
    blockCallTree = et.parse('C:\\pythonProject\\openess\\templates\\blokFunkcyjny.xml')
    blockCallRoot = blockCallTree.getroot()

    y = CodeXml(blockCallRoot)

    argList = ["zmienna1", "Data_block_1", "a1", "superZmienna", "Data_block_2", "y1", "Data_block_2" ,"y2", "DI_blokFunkcyjny_A10", "bardzo ważny komentarz"]
    y.completeArgs(argList)

    x.insertCode(y.code)
    i = i - 1

x.prepareCode()

fcTree.write('C:\\pythonProject\\openess\\imports\\funkcja2.xml')


