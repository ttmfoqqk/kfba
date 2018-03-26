<%
'/** 
' * nTPL ASP v1.0 (nTemplate ASP)
' * 
' * @filename : sky_template_class.asp
' * @version  : v 1.0 2003.08.18 14:56:22 
' *             v 0.9 2003.07.25 10:14:32
' *             v 0.5 2003.02.17 11:22:13
' * @author    : nsl (smpoem@magicn.com) (http://ndir.cyworld.com)
' * @copyright : Copyright �� 2003 nsl All rights reserved.
' *
' * @info
' *     VBScript 5.0 �̻���� ��밡���մϴ�.
' *     VBScript 5.5 �̻󿡼� ����ȭ �Ǿ����ϴ�.
' *
' */
class SkyTemplate

    '//  ���ø� ���丮�� (string)
    public tplDir

    '//^^^^^^^^^^^^^^^^^^^^^^^^ ���� Dictionary ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '// ���ø� ���ϸ� ���� (Dictionary)
    private fileList      
    
    '// ���ø� ���ϳ��� ���� (BLOCK �κи�) (Dictionary) 
    private fileValueList 

   '// ���ø� ���� ���� ����    (Dictionary)
    private lastValueList 

    '// BLOCK LOOP ó���� (Dictionary)
    private blockReplace  
    '//^^^^^^^^^^^^^^^^^^^^^^^^ /���� Dictionary ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


   '// ��� ��ũ "{" & varFirstMark & varSecondMark & "������}" => {$name}, {M$name} �� ...
   '// �����Ӱ� ������ �����ϴ�
    private varFirstMark  '// (string)
    private varSecondMark '// (string)

    private template  '// ����� ���ø� ���� (string)

   '// ������ ���� ã�� ������ ��� ������� ���� (boolean)
    public isBlockErrorCheck  

    private regex     '// RegExp object (object)
    private fso       '// ���� object (object)



    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ class new, nothing �ÿ� �ڵ� ���� �κ� ^^^^^^^^^^^^^^^^^^^^^

    '// Ŭ���� new(�ʱ�ȭ) �� �ڵ� ����
    private sub class_initialize()

    
        '// �⺻ ���� ��ũ���� {$������} 
        setVarMark "", "$"

        '// BLOCK �� ã�� ������ ��� ���� ��� ���� true(���), false(��¾���)
        setBlockErrorCheck(true)

        '// Dictionary, fileSystemObject ����
        set fileList      = createObject("Scripting.Dictionary") '// ���ϸ� ����
        set fileValueList = createObject("Scripting.Dictionary") '// BLOCK ���� ����
        set lastValueList = createObject("Scripting.Dictionary") '// ���� ġȯ�� ���� ����
        set blockReplace  = createObject("Scripting.Dictionary") '// BLOCK �κ� parse �Ҷ�

        set fso           = createObject("Scripting.FileSystemObject") '// ���ø� ���� ������ ���

        '// �⺻ ���ø� ���丮 ����
       '//setTplDir("tpl")
       tplDir = "tpl"

        '// ���Խ� ����
        set regex = new RegExp
        regex.ignoreCase = false '// ��ҹ��� ����
        regex.global     = true  '// ��ü���ڿ� �˻�

        template = ""

    end sub

    '// Ŭ���� nothing (set ntpl = nothing) �� ����
    private sub class_terminate()

        set fileList      = nothing
        set fileValueList = nothing
        set lastValueList = nothing
        set blockReplace  = nothing
        set fso           = nothing
        set regex         = nothing

    end sub
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /class new, nothing �ÿ� �ڵ� ���� �κ� ^^^^^^^^^^^^^^^^^^^^^

    '/*
    ' * ���ø� ���丮 �缳�� (�⺻ tpl)
    '
    ' * @param  string  newTplDir : ���� ������ ���ø� ���丮
    ' */
    public sub setTplDir(newTplDir)

        tplDir = newTplDir

        '// tpl ���丮�� �����ϴ��� �˻�
        if fso.folderExists(server.mapPath(tplDir)) = false then

            call errorMsg("'" & tplDir & "' �� ���ø� ���丮�� ã�� �� �����ϴ�!", false)
        end if

    end sub


    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ ��Ʋ�� ���� ��ũ ���� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '/**
    ' * ���ø� ������ ��ũ�� �����մϴ�.
    ' * �⺻ : {$������} �̳� {������}, {M$������} �� �پ��ϰ� �������� => ���ø� ���ϸ��� �ٸ��� ���������ϴ�
    ' *
    ' * @param  string  varFirstMark  : { + "�̺κ�" + varSecondMark + } => �ݵ�� ����
    ' * @param  string  varSecondMark : { + varFirstMark + "�̺κ�" + }  => �ݵ�� Ư������
    ' */
    public sub setVarMark(varFirstMarkTemp, varSecondMarkTemp)

         varFirstMark  = varFirstMarkTemp
         varSecondMark = varSecondMarkTemp
    end sub


    '/**
    ' * ġȯ�� ����ǥ���Ŀ��� ����� ���� ������ ����ϴ�.
    ' *
    ' * @param  string varName : ������
    ' *
    ' * @see tplParseBlock(), tplInBlockReset(), tplParse()
    ' */
    private function getPatterns(varName)

        getPatterns = "\{" & varFirstMark & "\" & varSecondMark & varName & "\}"

    end function


    '/**
    ' * ���� ������ ����� 
    ' * 
    ' * @param string varName : ������
    ' */
    private function getPatternsVar(varName)

        getPatternsVar = "{" & varFirstMark & varSecondMark & varName & "}"
    
    end function


    '/**
    ' * ���� ������ ���Ѵ�. (�� �����ϴ� �κп��� ���)
    ' *
    ' * @param  string  blockName : ���̸�
    ' *
    ' * @see getBlockNew()
    ' */
    private function getBlockPatterns(blockName) 

        getBlockPatterns = "<!--\s+BLOCK BEGIN\s+" & blockName & "\s+-->(.*)\n([\s\S.]*)<!--\s+BLOCK END\s+" & blockName & "\s+-->" '// \1
    end function

    '/**
    ' * include �� ���� ����
    ' *
    ' * @see setIncludeNewPatterns() 
    ' */
    private function getIncludePatterns

        '// [#]* -> nTPL PHP �������� ȣȯ�� ���ؼ�
        getIncludePatterns = "<!--\s+[#]*include file\s*=\s*[''""]([_a-zA-Z0-9_/.]+.[a-zA-Z]+)[''""]\s+-->"

    end function

    '/**
    ' * include �κ��� ���� �������� ġȯ�ϱ� ���� ����
    ' *
    ' * @param string include_file : include ���ϸ�
    ' *
    ' * @see getReadFile()
    ' */
    private function getIncludeReplacePatterns(include_file)

        getIncludeReplacePatterns = "<!--\s+[#]*include file\s*=\s*[''""](" & include_file & ")[''""]\s+-->"

    end function


    '/**
    ' * VBScript 5.5 �̸��� ��� include ���� ���� ����
    ' * <!-- #inclufe_file="test.thml" -->   ===> test.html
    ' *
    ' * @param string incfile : include �� ����
    ' *
    ' * @see getReadFile()
    ' */
    private function getIncludeFilename(incfile)

        dim temp

        '//  include �κ��� ���� �������� ġȯ�ϱ� ���� ���� ����
        regex.pattern = "<!--\s+[#]*include file\s*=\s*[''""]"
        temp = regex.replace(incfile, "")

        regex.pattern = "[''""]\s+-->"
        getIncludeFilename = regex.replace(temp, "")

    end function

    '/**
    ' * VBScript 5.5 �̸��� ��� block ���� ���� ����
    ' *
    ' * @param string block :
    ' * @param string blockNameTemp :
    ' *
    ' * @see setBlock()
    ' */
    private function getBlockNew(block, blockNameTemp) 

        dim temp

        regex.pattern = "<!--\s+BLOCK BEGIN\s+" & blockNameTemp & "\s+-->"
        temp = regex.replace(block, "")
        regex.pattern = "<!--\s+BLOCK END\s+" & blockNameTemp & "\s+-->"
        getBlockNew = regex.replace(temp, "")

    end function
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /��Ʋ�� ���� ��ũ ���� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



    '/*
    ' * ���ø����� ����� ������ �����մϴ�.
    ' *
    ' * @param  string, array  fileKey  : ���� Ű��
    ' * @param  string         fileName : ���ϸ�
    ' *
    ' * @example
    ' *    ntpl.setFile "HEADER", "header.html"  -> �ϳ��� ����
    ' *
    ' *    -> �ϳ��̻� ������ (�ݵ�� ��� ������� ����)
    ' *    ntpl.setFile  array(array("HEADER", "header.html"), array("BODY", "body.html")), ""
    ' */
    public function setFile(fileKey, fileName)

        dim fKeyTemp

        '// �迭�� �Ѿ�Դٸ�
        if  isArray(fileKey) AND not isNULL(fileKey) then

            '// �迭�� ����ŭ tpl ���� ����
            for each fKeyTemp in fileKey

                fileList.add fKeyTemp(0), fKeyTemp(1)
            next
        else

           '// tpl ���� �ϳ��� ����
           fileList.add fileKey, fileName
        end if

    end function


    '/*
    ' * ���ø����� ����� ���� ��� ������ ���� �ִ´�.
    ' *
    ' * @param  string, array  fileKey  : Ű ��
    ' * @param  string         contnet : ����
    ' *
    ' * @example
    ' *    ntpl.setFile "TEST", "����� ������ �����ִ´�"
    ' *
    ' */
    public function setFileAdd(fileKey, content)

        ' ����Ű�� �� �߰�
        fileList.add fileKey, ""

        fileValueList.add fileKey, Cstr(content)

    end function



    '/**
    ' * setFile() ���� ������ ���ø� ������ ���̼� => fileValueList["Ű��"] = "����" �迭�� �ִ´�
    ' *
    ' * @param  string   fileKey  : setFile() ���� ������ Ű���� �ϳ�
    ' * @param  boolean  isReturn : true(return), false
    ' *
    ' * @info
    ' *  - ������ �����鼭 INFO BLOCK �� �ּ����� ó���ؼ� �����Ѵ�
    ' */
    private function getReadFile(fileKey, isReturn)

        dim fileName, fp, fileContent
        dim match, matches
        dim include_file
        dim x


        '// �����̸� ����
        fileName = server.mapPath(tplDir & "/" & fileList.item(fileKey))

        '// ���ø� ������ �����ϸ�
        if fso.fileExists(fileName) then

            '// ���ø� ������ �д´�
            set fp = fso.openTextFile(fileName)

            '// ���� ����
            fileContent = fp.readAll


            '//^^^^^^^^^^^^^^^^^^^^^^^^^^ include ó�� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '// include �κ��� �������� ���� ����
            regex.pattern = getIncludePatterns()

            set matches = regex.execute(fileContent)

            '// include �κ��� ã�Ҵٸ�
            if matches.count > 0 then

                for each match in matches

                    '// VBScript 5.5 �̻��̸�, SubMatches �� 5.5 �̻���� ������
                    if ( isVbVer() ) then

                        '// include ���ϸ� ����
                        include_file = match.SubMatches(0)
                    else
                        '//            �������� ó��
                        include_file = getIncludeFilename(match.value)
                    end if

                    '// test code
                    '//response.write include_file & " - " & server.htmlEncode(match.value) & "<br>"

                    '//  include �κ��� ���� �������� ġȯ�ϱ� ���� ���� ����
                    regex.pattern = getIncludeReplacePatterns(include_file)

                    '// include �κ��� ���� ���ϳ������� ġȯ 
                    '//                          ( ���ϳ���, include ���� ���� )
                    fileContent = regex.replace(fileContent, getIncludeFile(include_file))

    	    	next

            end if
            '//^^^^^^^^^^^^^^^^^^^^^^^^^^ /include ó�� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        
            
            '// INFO BLOCK �� �����ϱ� ���� �κ� 
            regex.pattern = getBlockPatterns("INFO")

           '// 
            if isReturn then

                '// INFO BLOCK ������ return
                getReadFile = regex.replace(fileContent, "")
            else

                '// INFO BLOCK ������ ����
                fileValueList.add fileKey, regex.replace(fileContent, "")
            end if

            '// ������ �ݴ´�
            fp.close

        else
            '// ������ ã�� ������ ��� ����
            call errorMsg(fileName & " (" & fileKey & ") " & " �� ���ø� ������ ã�� �� �����ϴ�!", false)
        end if

     end function


    '/**
    ' * ���ø����� BLOCK �� ã�� ������ ��� ������ ����� ������ ����
    ' *
    ' * @param  boolean  isCheck : true(����üũ), false(üũ����)
    ' *
    ' * false �� �����ص� �������
    ' */
    public sub setBlockErrorCheck(isCheck) 

        isBlockErrorCheck = isCheck

    end sub

    '/*
    ' * BLOCK �κ��� ���ϰ� => fileValueList.item("������") = "���ѳ���" �ִ´�
    ' *
    ' * @param  string  fileKey   : setFile() ���� ������ ����Ű ���� �ϳ�
    ' * @param  string  blockName : ��Ʋ�� ���Ͽ� ������ BLOCK �̸�
    ' *
    ' * @example
    ' *    ntpl.setBlock "NTPL2", array("MEMBER_LOOP", "LIST", "NAME") -> �ϳ��̻� ����
    ' */
    public sub setBlock(fileKey, blockName)

        dim match, matches, blockVarName
        dim eMsg
        dim matchValueTemp
        dim blockNameTemp

        '// ���ø� ������ ���� �ʾҴٸ� 
        if (fileValueList.exists(fileKey) = false) then

            '// ������ �д´�
            call getReadFile(fileKey, false)
        end if

        dim rKeyTemp


        for each blockNameTemp in blockName

           '// BLOCK ������ ���� ����
            regex.pattern = getBlockPatterns(blockNameTemp)

            set matches = regex.execute(fileValueList.item(fileKey))

            '// BLOCK �� ã�Ҵٸ�
            if matches.count > 0 then

                '// BLOCK => {������} ���� ġȯ�ϱ� ���� {������} ����
                blockVarName = getPatternsVar(blockNameTemp)

                for each match in matches

                    '// VBScript 5.5 �̻��̸�, SubMatches �� 5.5 �̻���� ������
                    if ( isVbVer() ) then

                        '// ������ BLOCK �κ� ����
                        fileValueList.item(blockNameTemp) =  match.SubMatches(1) '//match.value
                    else

                        fileValueList.item(blockNameTemp) =  getBlockNew(match.value, blockNameTemp) '// ���� ���簡 ����
                    end if

                    if not isNull(blockVarName) then

                        '// BLOCK �κ� => "" �ǰ� �׿� �κ� ����
                        fileValueList.item(fileKey) = regex.replace(fileValueList.item(fileKey), Cstr(blockVarName))
                    else

                        '// BLOCK �κ� => "" �ǰ� �׿� �κ� ����
                        fileValueList.item(fileKey) = regex.replace(fileValueList.item(fileKey), "")
                    end if

                    ' test code
                    'response.write "<p><font color=blue>ã�� : [[" &  server.htmlencode(match.SubMatches(1)) & "]]</font><br>"
    	    	next

           else
            '// BLOCK �� ã�� ������ ��

                '// BLOCK �� ã�� �������� ���� ����̸�
                if isBlockErrorCheck then

                    '// ������ ����Ѵ�
                    eMsg = tplDir & "/" & fileList.item(fileKey) & " �� ���Ͽ���<br>" & _
                           " '" & blockNameTemp & "' �� BLOCK ���� ã�� �� �����ϴ�!"
                    call errorMsg(eMsg, false)
                end if
    
           end if

        next

    end sub

    
    '/**
    ' * BLOCK �κ��� replace
    ' *
    ' * @param  string, array  rKey  : Ű
    ' * @param  string  rItem : ��
    ' *
    ' * @example
    ' *   ntpl.setBlockReplace "name", "test" -> �ϳ��� ����
    ' *   ntpl.setBlockReplace array(array("name", "test"), array("userid", "testid") ), "" -> �ϳ��̻� ����
    ' */
    public sub setBlockReplace(rKey, rItem)

        dim rKeyTemp

        '// �迭�� �Ѿ�Դٸ�
        if  isArray(rKey) AND not isNULL(rKey) then

            for each rKeyTemp in rKey

                blockReplace.add rKeyTemp(0), rKeyTemp(1)
            next
        else

            blockReplace.add rKey, rItem
        end if
    end sub


    '/**
    ' * BLOCK LOOP parse �κ�
    ' *
    ' * @param  string  fileKey  : setFile() ���� ������ ������ Ű������ �ϳ�
    ' *
    ' * @example
    ' *    ntpl.setBlockReplace "name", "�׽�Ʈ"  --> �� ���� ����
    ' *    nptl.tplParseBlock("LIST")
    ' */
    public sub tplParseBlock(fileKey)

        dim lKey, strTemp

        strTemp = fileValueList.item(fileKey)

        '// {$����} => '����' ���� �ٲ۴�.
        for each lKey in blockReplace


            '// pattern ����
            regex.pattern = getPatterns(lKey)

            if not isNull(blockReplace.item(lKey)) then

                strTemp = regex.replace(strTemp, Cstr(blockReplace.item(lKey)))
            else

                strTemp = regex.replace(strTemp, "")
			end if
        next

        lastValueList.item(fileKey) = lastValueList.item(fileKey) & strTemp

        '// ��� ������
        blockReplace.removeAll

    end sub


    '/*
    ' * ���� ��� ������ ���� �����Ѵ�
    ' *
    ' * @param  string rKey  : Ű��
    ' * @param  string rItem : ����
    ' *
    ' * @example
    ' *    ntpl.setLastValue "Ű��", "����"
    ' */
    public function setLastValue(rKey, rItem)

        lastValueList.item(rKey) = rItem

    end function


    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ �� ����, ���� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '/**
    ' * �����ߴ� ���� -> �����ߴ� ������ �ٽ� �ִ´�
    ' *
    ' * @param  string  blockName : ���̸�
    ' * @param  string  resetValue : ���� ��
    ' */
    public sub tplBlockReset(blockName, resetValue)

        if resetValue = "" then
            lastValueList.item(blockName) = fileValueList.item(blockName)
        else
            lastValueList.remove(blockName)
        end if

    end sub

    '/**
    ' * �����ߴ� ���κ� {TEST}�� �����Ѵ�.
    ' *
    ' * @param  string blockName : ���̸�
    ' */
    public sub tplBlockDel(blockName)

        lastValueList.item(blockName) = ""
        '//lastValueList.remove(blockName)

    end sub

    '/**
    ' * �����ߴ� ��{TEST}���� ��{TEST_LOOP}�� => �����ߴ� ������ �ٽ� �ִ´�
    ' *
    ' * @param  string blockName : ���̸�
    ' * @param  string blockNameSub : �ٲ� ���̸�
    ' * @param  boolean isNULL      : true(������)
    ' */
    public sub tplInBlockReset(blockName, blockNameSub, isDel)

        dim replaces

        '// pattern ����
         regex.pattern = getPatterns(blockNameSub)

         '// �����̸�
         if isDel then
             replaces = ""
         else
             replaces = fileValueList.item(blockNameSub)
         end if
         
         fileValueList.item(blockName) = regex.replace(fileValueList.item(blockName), replaces)

         '// ������
         tplBlockDel(blockNameSub)

    end sub
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /�� ����, ���� ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '/**
    ' * lastValueList["������"] = '���� �ٲ� ����' ���� �ִ´�
    ' *
    ' * @param  string, array  $varName : ������
    ' * @param  string         $varValue : �ٲܳ���
    ' *
    ' * @example
    ' *    ntpl.tplAssign "name", "�׽�Ʈ"  -> �ϳ��� ����
    ' *    ntpl.tplAssing array(array("name", "�׽�Ʈ"), array("userid", "test")), "" -> �ϳ��̻� ������
    ' */
    public sub tplAssign(varName, varValue)
        
        dim vNameTemp

        '// �迭�� �Ѿ�Դٸ�
        if  isArray(varName) AND not isNULL(varName) then

            for each vNameTemp in varName
                lastValueList.add vNameTemp(0), vNameTemp(1)
            next
        else

            lastValueList.add varName, varValue
        end if

    end sub



    '/**
    ' * ��� �� parse
    ' *
    ' * template �� ���� ������ ��´�
    ' */
    public sub tplParse() 

        dim keyValue, replaces, patterns
        dim flKey
        dim strAll
        dim lKey

        '// setFile() ���� ������ ��� ���ø� ������ ��� ������ ����
        for each flKey in fileList

            '// fileList(����Ű) �� ������ => fileValueList(����Ű) �� �ִٸ�
            '// setBlock() ���� BLOCK ������ ������ �̹� �о���
            if( fileValueList.exists(flKey) ) then
                
                strAll = strAll & fileValueList.item(flKey)
            else

                strAll = strAll & getReadFile(flKey, true)
            end if

        next

        '// �ӵ� ������ �κ�
        for each lKey in lastValueList

            '// pattern ����
            regex.pattern = getPatterns(lKey)

            ' test code
            'response.write "<br>pattern : " & regex.pattern
            'response.write " - replace : "  & Server.HTMLEncode(lastValueList.item(lKey))
		
            if not isNull( lastValueList.item(lKey) ) then

                strAll = regex.replace(strAll, Cstr( lastValueList.item(lKey)) )
            else

                strAll = regex.replace(strAll, "" )
            end if

        next

        template = strAll
        set strAll = nothing
         
    end sub


    ' ���ø��� parse �� ���� ���� ���
    public sub tplPrint()

        response.write template

    end sub


    ' ���ø� ������ ���Ѵ�
    public function getTplContent()

       getTplContent = template
    end function

    '/**
    ' * include �� ���� ������ ���Ѵ�
    ' *
    ' * @param  string  file : include �� ����
    ' */
    private function getIncludeFile(include_file) 

        dim  filename, fp

        '// ���� ��η�
        filename = Server.MapPath(include_file)

        '// ������ ã�� ���ߴٸ�
        if fso.fileExists(filename) = false then
            
            call errorMsg(filename & " �� ������ ã�� �� �����ϴ�. (include)", false)

        else
        '// ������ ã�Ҵٸ�

            '// ���ø� ������ �д´�
            set fp = fso.openTextFile(filename)

            '// ���ϳ��� ����
            getIncludeFile = fp.readAll
  
            '// ������ �ݴ´�
            fp.close

        end if

    end function


    '/*
    ' * ���� �޼��� ���
    ' *
    ' * @param  string   msg : ����� ���� �޼���
    ' * @param  boolean  isEnd : true(����), false
    ' */
    private sub errorMsg(msg, isEnd)

        response.write "<p><font color=red style='font-style:9pt'>* nslTemplate Error : " & msg & "</font><p>"

        if (isEnd) then
            response.end
        end if
    end sub
    
    '// ScriptEngine ���� üũ 5.5 �̻󿡼��� ��밡���� ���(���� ������)
    private function getScriptEngineInfo

        dim SEVer
  
       '//               5                 .             6
       SEVer = ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion 

       if ScriptEngine <> "VBScript" OR SEVer < 5.5 then

           SEVer = ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion 

           call errorMsg("VBScript 5.5 �̻󿡼��� ����� �� �ֽ��ϴ�! (���� ���� ���� : <font color=blue>" & SEVer & "</font> �Դϴ�.)", true)
       end if
   
    end Function

    '// VBScript 5.5 �ΰ�
    private function isVbVer()

       dim SEVer

       '//               5                 .             6
       SEVer = ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion

       if SEVer >= 5.5 then
           isVbVer = true
       else
           isVbVer = false
       end if

    end function

end class
%>