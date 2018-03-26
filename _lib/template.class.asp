<%
'/** 
' * nTPL ASP v1.0 (nTemplate ASP)
' * 
' * @filename : sky_template_class.asp
' * @version  : v 1.0 2003.08.18 14:56:22 
' *             v 0.9 2003.07.25 10:14:32
' *             v 0.5 2003.02.17 11:22:13
' * @author    : nsl (smpoem@magicn.com) (http://ndir.cyworld.com)
' * @copyright : Copyright ⓒ 2003 nsl All rights reserved.
' *
' * @info
' *     VBScript 5.0 이상부터 사용가능합니다.
' *     VBScript 5.5 이상에서 최적화 되었습니다.
' *
' */
class SkyTemplate

    '//  템플릿 디렉토리명 (string)
    public tplDir

    '//^^^^^^^^^^^^^^^^^^^^^^^^ 저장 Dictionary ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '// 템플릿 파일명 저장 (Dictionary)
    private fileList      
    
    '// 템플릿 파일내용 저장 (BLOCK 부분만) (Dictionary) 
    private fileValueList 

   '// 템플릿 최종 내용 저장    (Dictionary)
    private lastValueList 

    '// BLOCK LOOP 처리시 (Dictionary)
    private blockReplace  
    '//^^^^^^^^^^^^^^^^^^^^^^^^ /저장 Dictionary ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


   '// 사용 마크 "{" & varFirstMark & varSecondMark & "변수명}" => {$name}, {M$name} 등 ...
   '// 자유롭게 설정이 가능하다
    private varFirstMark  '// (string)
    private varSecondMark '// (string)

    private template  '// 출력할 템플릿 내용 (string)

   '// 지정한 블럭을 찾지 못했을 경우 에러출력 유무 (boolean)
    public isBlockErrorCheck  

    private regex     '// RegExp object (object)
    private fso       '// 파일 object (object)



    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ class new, nothing 시에 자동 실행 부분 ^^^^^^^^^^^^^^^^^^^^^

    '// 클래스 new(초기화) 시 자동 실행
    private sub class_initialize()

    
        '// 기본 변수 마크설정 {$변수명} 
        setVarMark "", "$"

        '// BLOCK 을 찾지 못했을 경우 에러 출력 유무 true(출력), false(출력안함)
        setBlockErrorCheck(true)

        '// Dictionary, fileSystemObject 생성
        set fileList      = createObject("Scripting.Dictionary") '// 파일명 저장
        set fileValueList = createObject("Scripting.Dictionary") '// BLOCK 추출 저장
        set lastValueList = createObject("Scripting.Dictionary") '// 최종 치환할 내용 저장
        set blockReplace  = createObject("Scripting.Dictionary") '// BLOCK 부분 parse 할때

        set fso           = createObject("Scripting.FileSystemObject") '// 템플릿 파일 읽을때 사용

        '// 기본 템플릿 디렉토리 설정
       '//setTplDir("tpl")
       tplDir = "tpl"

        '// 정규식 설정
        set regex = new RegExp
        regex.ignoreCase = false '// 대소문자 구분
        regex.global     = true  '// 전체문자열 검색

        template = ""

    end sub

    '// 클래스 nothing (set ntpl = nothing) 시 실행
    private sub class_terminate()

        set fileList      = nothing
        set fileValueList = nothing
        set lastValueList = nothing
        set blockReplace  = nothing
        set fso           = nothing
        set regex         = nothing

    end sub
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /class new, nothing 시에 자동 실행 부분 ^^^^^^^^^^^^^^^^^^^^^

    '/*
    ' * 템플릿 디렉토리 재설정 (기본 tpl)
    '
    ' * @param  string  newTplDir : 새로 설정할 템플릿 디렉토리
    ' */
    public sub setTplDir(newTplDir)

        tplDir = newTplDir

        '// tpl 디렉토리가 존재하는지 검사
        if fso.folderExists(server.mapPath(tplDir)) = false then

            call errorMsg("'" & tplDir & "' 이 템플릿 디렉토리를 찾을 수 없습니다!", false)
        end if

    end sub


    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 템틀릿 변수 마크 설정 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '/**
    ' * 템플릿 변수의 마크를 설정합니다.
    ' * 기본 : {$변수명} 이나 {변수명}, {M$변수명} 등 다양하게 설정가능 => 템플릿 파일마다 다르게 설정가능하다
    ' *
    ' * @param  string  varFirstMark  : { + "이부분" + varSecondMark + } => 반드시 문자
    ' * @param  string  varSecondMark : { + varFirstMark + "이부분" + }  => 반드시 특수문자
    ' */
    public sub setVarMark(varFirstMarkTemp, varSecondMarkTemp)

         varFirstMark  = varFirstMarkTemp
         varSecondMark = varSecondMarkTemp
    end sub


    '/**
    ' * 치환시 정규표현식에서 사용할 변수 패턴을 만듭니다.
    ' *
    ' * @param  string varName : 변수명
    ' *
    ' * @see tplParseBlock(), tplInBlockReset(), tplParse()
    ' */
    private function getPatterns(varName)

        getPatterns = "\{" & varFirstMark & "\" & varSecondMark & varName & "\}"

    end function


    '/**
    ' * 패턴 변수명 만들기 
    ' * 
    ' * @param string varName : 변수명
    ' */
    private function getPatternsVar(varName)

        getPatternsVar = "{" & varFirstMark & varSecondMark & varName & "}"
    
    end function


    '/**
    ' * 블럭의 패턴을 구한다. (블럭 추출하는 부분에서 사용)
    ' *
    ' * @param  string  blockName : 블럭이름
    ' *
    ' * @see getBlockNew()
    ' */
    private function getBlockPatterns(blockName) 

        getBlockPatterns = "<!--\s+BLOCK BEGIN\s+" & blockName & "\s+-->(.*)\n([\s\S.]*)<!--\s+BLOCK END\s+" & blockName & "\s+-->" '// \1
    end function

    '/**
    ' * include 의 패턴 구함
    ' *
    ' * @see setIncludeNewPatterns() 
    ' */
    private function getIncludePatterns

        '// [#]* -> nTPL PHP 버전과의 호환을 위해서
        getIncludePatterns = "<!--\s+[#]*include file\s*=\s*[''""]([_a-zA-Z0-9_/.]+.[a-zA-Z]+)[''""]\s+-->"

    end function

    '/**
    ' * include 부분을 실제 내용으로 치환하기 위한 패턴
    ' *
    ' * @param string include_file : include 파일명
    ' *
    ' * @see getReadFile()
    ' */
    private function getIncludeReplacePatterns(include_file)

        getIncludeReplacePatterns = "<!--\s+[#]*include file\s*=\s*[''""](" & include_file & ")[''""]\s+-->"

    end function


    '/**
    ' * VBScript 5.5 미만일 경우 include 패턴 잔재 삭제
    ' * <!-- #inclufe_file="test.thml" -->   ===> test.html
    ' *
    ' * @param string incfile : include 문 패턴
    ' *
    ' * @see getReadFile()
    ' */
    private function getIncludeFilename(incfile)

        dim temp

        '//  include 부분을 실제 내용으로 치환하기 위한 패턴 구함
        regex.pattern = "<!--\s+[#]*include file\s*=\s*[''""]"
        temp = regex.replace(incfile, "")

        regex.pattern = "[''""]\s+-->"
        getIncludeFilename = regex.replace(temp, "")

    end function

    '/**
    ' * VBScript 5.5 미만일 경우 block 패턴 잔재 삭제
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
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /템틀릿 변수 마크 설정 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^



    '/*
    ' * 템플릿으로 사용할 파일을 설정합니다.
    ' *
    ' * @param  string, array  fileKey  : 파일 키값
    ' * @param  string         fileName : 파일명
    ' *
    ' * @example
    ' *    ntpl.setFile "HEADER", "header.html"  -> 하나만 설정
    ' *
    ' *    -> 하나이상 설정시 (반드시 출력 순서대로 설정)
    ' *    ntpl.setFile  array(array("HEADER", "header.html"), array("BODY", "body.html")), ""
    ' */
    public function setFile(fileKey, fileName)

        dim fKeyTemp

        '// 배열로 넘어왔다면
        if  isArray(fileKey) AND not isNULL(fileKey) then

            '// 배열의 수만큼 tpl 파일 설정
            for each fKeyTemp in fileKey

                fileList.add fKeyTemp(0), fKeyTemp(1)
            next
        else

           '// tpl 파일 하나만 설정
           fileList.add fileKey, fileName
        end if

    end function


    '/*
    ' * 템플릿으로 사용할 파일 대신 내용을 직접 넣는다.
    ' *
    ' * @param  string, array  fileKey  : 키 값
    ' * @param  string         contnet : 내용
    ' *
    ' * @example
    ' *    ntpl.setFile "TEST", "출력할 내용을 직접넣는다"
    ' *
    ' */
    public function setFileAdd(fileKey, content)

        ' 파일키에 빈값 추가
        fileList.add fileKey, ""

        fileValueList.add fileKey, Cstr(content)

    end function



    '/**
    ' * setFile() 에서 설정한 템플릿 파일을 읽이서 => fileValueList["키값"] = "내용" 배열에 넣는다
    ' *
    ' * @param  string   fileKey  : setFile() 에서 설정한 키값중 하나
    ' * @param  boolean  isReturn : true(return), false
    ' *
    ' * @info
    ' *  - 파일을 읽으면서 INFO BLOCK 은 주석으로 처리해서 삭제한다
    ' */
    private function getReadFile(fileKey, isReturn)

        dim fileName, fp, fileContent
        dim match, matches
        dim include_file
        dim x


        '// 파일이름 조합
        fileName = server.mapPath(tplDir & "/" & fileList.item(fileKey))

        '// 템플릿 파일이 존재하면
        if fso.fileExists(fileName) then

            '// 템플릿 파일을 읽는다
            set fp = fso.openTextFile(fileName)

            '// 파일 내용
            fileContent = fp.readAll


            '//^^^^^^^^^^^^^^^^^^^^^^^^^^ include 처리 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            '// include 부분을 추출하지 위한 패턴
            regex.pattern = getIncludePatterns()

            set matches = regex.execute(fileContent)

            '// include 부분을 찾았다면
            if matches.count > 0 then

                for each match in matches

                    '// VBScript 5.5 이상이면, SubMatches 가 5.5 이상부터 지원됨
                    if ( isVbVer() ) then

                        '// include 파일명 구함
                        include_file = match.SubMatches(0)
                    else
                        '//            패턴잔재 처리
                        include_file = getIncludeFilename(match.value)
                    end if

                    '// test code
                    '//response.write include_file & " - " & server.htmlEncode(match.value) & "<br>"

                    '//  include 부분을 실제 내용으로 치환하기 위한 패턴 구함
                    regex.pattern = getIncludeReplacePatterns(include_file)

                    '// include 부분을 실제 파일내용으로 치환 
                    '//                          ( 파일내용, include 파일 내용 )
                    fileContent = regex.replace(fileContent, getIncludeFile(include_file))

    	    	next

            end if
            '//^^^^^^^^^^^^^^^^^^^^^^^^^^ /include 처리 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        
            
            '// INFO BLOCK 을 제거하기 위한 부분 
            regex.pattern = getBlockPatterns("INFO")

           '// 
            if isReturn then

                '// INFO BLOCK 삭제후 return
                getReadFile = regex.replace(fileContent, "")
            else

                '// INFO BLOCK 삭제후 저장
                fileValueList.add fileKey, regex.replace(fileContent, "")
            end if

            '// 파일을 닫는다
            fp.close

        else
            '// 파일을 찾지 못했을 경우 에러
            call errorMsg(fileName & " (" & fileKey & ") " & " 이 템플릿 파일을 찾을 수 없습니다!", false)
        end if

     end function


    '/**
    ' * 템플릿에서 BLOCK 을 찾지 못했을 경우 에러를 출력할 것인지 설정
    ' *
    ' * @param  boolean  isCheck : true(에러체크), false(체크안함)
    ' *
    ' * false 로 설정해도 관계없음
    ' */
    public sub setBlockErrorCheck(isCheck) 

        isBlockErrorCheck = isCheck

    end sub

    '/*
    ' * BLOCK 부분을 구하고 => fileValueList.item("변수명") = "구한내용" 넣는다
    ' *
    ' * @param  string  fileKey   : setFile() 에서 설정한 파일키 중의 하나
    ' * @param  string  blockName : 텡틀릿 파일에 지정된 BLOCK 이름
    ' *
    ' * @example
    ' *    ntpl.setBlock "NTPL2", array("MEMBER_LOOP", "LIST", "NAME") -> 하나이상 설정
    ' */
    public sub setBlock(fileKey, blockName)

        dim match, matches, blockVarName
        dim eMsg
        dim matchValueTemp
        dim blockNameTemp

        '// 템플릿 파일을 읽지 않았다면 
        if (fileValueList.exists(fileKey) = false) then

            '// 파일을 읽는다
            call getReadFile(fileKey, false)
        end if

        dim rKeyTemp


        for each blockNameTemp in blockName

           '// BLOCK 추출을 위한 패턴
            regex.pattern = getBlockPatterns(blockNameTemp)

            set matches = regex.execute(fileValueList.item(fileKey))

            '// BLOCK 를 찾았다면
            if matches.count > 0 then

                '// BLOCK => {변수명} 으로 치환하기 위한 {변수명} 구함
                blockVarName = getPatternsVar(blockNameTemp)

                for each match in matches

                    '// VBScript 5.5 이상이면, SubMatches 가 5.5 이상부터 지원됨
                    if ( isVbVer() ) then

                        '// 추출한 BLOCK 부분 저장
                        fileValueList.item(blockNameTemp) =  match.SubMatches(1) '//match.value
                    else

                        fileValueList.item(blockNameTemp) =  getBlockNew(match.value, blockNameTemp) '// 패턴 잔재가 남음
                    end if

                    if not isNull(blockVarName) then

                        '// BLOCK 부분 => "" 되고 그외 부분 저장
                        fileValueList.item(fileKey) = regex.replace(fileValueList.item(fileKey), Cstr(blockVarName))
                    else

                        '// BLOCK 부분 => "" 되고 그외 부분 저장
                        fileValueList.item(fileKey) = regex.replace(fileValueList.item(fileKey), "")
                    end if

                    ' test code
                    'response.write "<p><font color=blue>찾다 : [[" &  server.htmlencode(match.SubMatches(1)) & "]]</font><br>"
    	    	next

           else
            '// BLOCK 을 찾지 못했을 때

                '// BLOCK 를 찾지 못했을때 에러 출력이면
                if isBlockErrorCheck then

                    '// 에러를 출력한다
                    eMsg = tplDir & "/" & fileList.item(fileKey) & " 이 파일에서<br>" & _
                           " '" & blockNameTemp & "' 이 BLOCK 문을 찾을 수 없습니다!"
                    call errorMsg(eMsg, false)
                end if
    
           end if

        next

    end sub

    
    '/**
    ' * BLOCK 부분의 replace
    ' *
    ' * @param  string, array  rKey  : 키
    ' * @param  string  rItem : 값
    ' *
    ' * @example
    ' *   ntpl.setBlockReplace "name", "test" -> 하나만 설정
    ' *   ntpl.setBlockReplace array(array("name", "test"), array("userid", "testid") ), "" -> 하나이상 설정
    ' */
    public sub setBlockReplace(rKey, rItem)

        dim rKeyTemp

        '// 배열로 넘어왔다면
        if  isArray(rKey) AND not isNULL(rKey) then

            for each rKeyTemp in rKey

                blockReplace.add rKeyTemp(0), rKeyTemp(1)
            next
        else

            blockReplace.add rKey, rItem
        end if
    end sub


    '/**
    ' * BLOCK LOOP parse 부분
    ' *
    ' * @param  string  fileKey  : setFile() 에설 설정한 파일의 키값중의 하나
    ' *
    ' * @example
    ' *    ntpl.setBlockReplace "name", "테스트"  --> 꼭 먼저 설정
    ' *    nptl.tplParseBlock("LIST")
    ' */
    public sub tplParseBlock(fileKey)

        dim lKey, strTemp

        strTemp = fileValueList.item(fileKey)

        '// {$변수} => '내용' 으로 바꾼다.
        for each lKey in blockReplace


            '// pattern 만듬
            regex.pattern = getPatterns(lKey)

            if not isNull(blockReplace.item(lKey)) then

                strTemp = regex.replace(strTemp, Cstr(blockReplace.item(lKey)))
            else

                strTemp = regex.replace(strTemp, "")
			end if
        next

        lastValueList.item(fileKey) = lastValueList.item(fileKey) & strTemp

        '// 모두 삭제함
        blockReplace.removeAll

    end sub


    '/*
    ' * 최종 출력 내용을 직접 설정한다
    ' *
    ' * @param  string rKey  : 키값
    ' * @param  string rItem : 내용
    ' *
    ' * @example
    ' *    ntpl.setLastValue "키값", "내용"
    ' */
    public function setLastValue(rKey, rItem)

        lastValueList.item(rKey) = rItem

    end function


    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ 블럭 수정, 삭제 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '/**
    ' * 추출했던 블럭에 -> 추출했던 내용을 다시 넣는다
    ' *
    ' * @param  string  blockName : 블럭이름
    ' * @param  string  resetValue : 넣을 값
    ' */
    public sub tplBlockReset(blockName, resetValue)

        if resetValue = "" then
            lastValueList.item(blockName) = fileValueList.item(blockName)
        else
            lastValueList.remove(blockName)
        end if

    end sub

    '/**
    ' * 추출했던 블럭부분 {TEST}을 삭제한다.
    ' *
    ' * @param  string blockName : 블럭이름
    ' */
    public sub tplBlockDel(blockName)

        lastValueList.item(blockName) = ""
        '//lastValueList.remove(blockName)

    end sub

    '/**
    ' * 추출했던 블럭{TEST}안의 블럭{TEST_LOOP}에 => 추출했던 내용을 다시 넣는다
    ' *
    ' * @param  string blockName : 블럭이름
    ' * @param  string blockNameSub : 바꿀 블럭이름
    ' * @param  boolean isNULL      : true(블럭삭제)
    ' */
    public sub tplInBlockReset(blockName, blockNameSub, isDel)

        dim replaces

        '// pattern 만듬
         regex.pattern = getPatterns(blockNameSub)

         '// 삭제이면
         if isDel then
             replaces = ""
         else
             replaces = fileValueList.item(blockNameSub)
         end if
         
         fileValueList.item(blockName) = regex.replace(fileValueList.item(blockName), replaces)

         '// 블럭삭제
         tplBlockDel(blockNameSub)

    end sub
    '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ /블럭 수정, 삭제 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    '/**
    ' * lastValueList["변수명"] = '실제 바꿀 내용' 으로 넣는다
    ' *
    ' * @param  string, array  $varName : 변수명
    ' * @param  string         $varValue : 바꿀내용
    ' *
    ' * @example
    ' *    ntpl.tplAssign "name", "테스트"  -> 하나만 설정
    ' *    ntpl.tplAssing array(array("name", "테스트"), array("userid", "test")), "" -> 하나이상 설정시
    ' */
    public sub tplAssign(varName, varValue)
        
        dim vNameTemp

        '// 배열로 넘어왔다면
        if  isArray(varName) AND not isNULL(varName) then

            for each vNameTemp in varName
                lastValueList.add vNameTemp(0), vNameTemp(1)
            next
        else

            lastValueList.add varName, varValue
        end if

    end sub



    '/**
    ' * 모든 것 parse
    ' *
    ' * template 에 최종 내용을 담는다
    ' */
    public sub tplParse() 

        dim keyValue, replaces, patterns
        dim flKey
        dim strAll
        dim lKey

        '// setFile() 에서 설정한 모든 템플릿 파일의 모든 내용을 구함
        for each flKey in fileList

            '// fileList(파일키) 의 내용이 => fileValueList(파일키) 에 있다면
            '// setBlock() 에서 BLOCK 설정한 파일은 이미 읽었다
            if( fileValueList.exists(flKey) ) then
                
                strAll = strAll & fileValueList.item(flKey)
            else

                strAll = strAll & getReadFile(flKey, true)
            end if

        next

        '// 속도 개선할 부분
        for each lKey in lastValueList

            '// pattern 만듬
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


    ' 템플릿의 parse 된 최종 내용 출력
    public sub tplPrint()

        response.write template

    end sub


    ' 템플릿 내용을 구한다
    public function getTplContent()

       getTplContent = template
    end function

    '/**
    ' * include 할 파일 내용을 구한다
    ' *
    ' * @param  string  file : include 할 파일
    ' */
    private function getIncludeFile(include_file) 

        dim  filename, fp

        '// 절대 경로로
        filename = Server.MapPath(include_file)

        '// 파일을 찾지 못했다면
        if fso.fileExists(filename) = false then
            
            call errorMsg(filename & " 이 파일을 찾을 수 없습니다. (include)", false)

        else
        '// 파일을 찾았다면

            '// 템플릿 파일을 읽는다
            set fp = fso.openTextFile(filename)

            '// 파일내용 리턴
            getIncludeFile = fp.readAll
  
            '// 파일을 닫는다
            fp.close

        end if

    end function


    '/*
    ' * 에러 메세지 출력
    ' *
    ' * @param  string   msg : 출력할 에러 메세지
    ' * @param  boolean  isEnd : true(종료), false
    ' */
    private sub errorMsg(msg, isEnd)

        response.write "<p><font color=red style='font-style:9pt'>* nslTemplate Error : " & msg & "</font><p>"

        if (isEnd) then
            response.end
        end if
    end sub
    
    '// ScriptEngine 버전 체크 5.5 이상에서만 사용가능할 경우(현재 사용안함)
    private function getScriptEngineInfo

        dim SEVer
  
       '//               5                 .             6
       SEVer = ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion 

       if ScriptEngine <> "VBScript" OR SEVer < 5.5 then

           SEVer = ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion 

           call errorMsg("VBScript 5.5 이상에서만 사용할 수 있습니다! (현재 서버 버전 : <font color=blue>" & SEVer & "</font> 입니다.)", true)
       end if
   
    end Function

    '// VBScript 5.5 인가
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