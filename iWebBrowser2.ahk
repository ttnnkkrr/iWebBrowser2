

class iWebBrowser2 {

    __new( ByRef sURL = "", ByRef sTitle = "", ByRef iHWND = "",  ByRef sHTML = "" , bVisible = true ) 
        { ;; returns an iwebbrowser2 object
        if !(this.pwb := this.oIE_get(sTitle,iHWND,sURL,sHTML))
            {
            this.pwb := this.oIE_new(sURL,sTitle,iHWND,sHTML,bVisible)
            this.pwb.Navigate(sURL)
            }
            
        return this
        }
    
    oIE_new(ByRef sURL = "about:blank", ByRef sTitle = "", ByRef iHWND = "", ByRef sHTML = "", bVisible = true)
        {
        this.isInstalled()
        this.pwb := ComObjCreate("internetexplorer.application")
        this.pwb.Visible := bVisible
        if sURL
            this.pwb.Navigate(sURL)
        return this.pwb
        }

    oIE_get( ByRef sTitle = "", ByRef iHWND = "", ByRef sURL = "", ByRef sHTML = "" )
        {
        this.isInstalled()
        ;~ this function is pointless if no instance of IE is open
        ;~ one edit you might make is to have this function open IE and maybe go to the home page
        if ( !winexist( "ahk_class IEFrame" ) )
            return false
        
        if sTitle
            this.clean_IE_Title( sTitle ) 
        ;; ok this function should look at all the existing IE instances and build a reference object
        ; List all open Explorer and Internet Explorer windows:
        oIE := Object()
        matches := 0
        
        for window,k in ComObjCreate("Shell.Application").Windows
            if ( "Internet Explorer" = window.Name)
                {
                possiblematch := true

                try pdoc := window.document
                Catch, e
                    while window.busy 
                        Sleep, 500

                if !window.document
                    Continue

                if ( possiblematch && sTitle && !instr( pdoc.title, sTitle ) )
                    possiblematch := false
                
                if ( possiblematch && sHTML && !instr( pdoc.documentelement.outerhtml, sHTML ) )
                    possiblematch := false
                
                if ( possiblematch && sURL && !instr( pdoc.url, sURL ) )
                    possiblematch := false
                
                if ( possiblematch && iHWND > 0 && window.HWND != iHWND )
                    possiblematch := false		
                    
                if ( possiblematch )
                    {
                    ;~ windowsList .= k " => " ( clipboard := window.FullName ) " :: " pdoc.title " :: " pdoc.url "`n"
                    matches++
                    sTitle := pdoc.title
                    sURL := pdoc.url
                    iHWND := window.HWND
                    sHTML := pdoc.documentelement.outerhtml
                    oIE := window
                    }
                ObjRelease( pdoc )
                }
                
        if ( matches > 1 )
            {
            MsgBox, 4112, Too many Matches ,  Please modify your criteria or close some tabs/windows and retry
            return false
            }
        
        return this.pwb := oIE
        }

    FindbyText(needle)
        { ;; returns the element with the text in it  
        try rng:=this.pdoc().body.createTextRange()
        try rng.findText(needle)
		return try rng.parentElement()
	    }

    Activate(pwb) 
        { 
        DllCall("LoadLibrary", "str", "oleacc.dll") 
        HWND:=pwb.HWND
        DetectHiddenWindows, On 
        WinActivate,% "ahk_id " HWND
        WinWaitActive,% "ahk_id " HWND,,5
        ControlGet, hTabBand, hWnd,, TabBandClass1, ahk_class IEFrame
        ControlGet, hTabUI  , hWnd,, DirectUIHWND1, ahk_id %hTabBand% 
        
        VarSetCapacity(CLSID, 16)
        nSize=38
        wString := sString := "{618736E0-3C3D-11CF-810C-00AA00389B71}"
        if(nSize = "")
            nSize:=DllCall("kernel32\MultiByteToWideChar", "Uint", 0, "Uint", 0, "Uint", &sString, "int", -1, "Uint", 0, "int", 0)
        VarSetCapacity(wString, nSize * 2 + 1)
        DllCall("kernel32\MultiByteToWideChar", "Uint", 0, "Uint", 0, "Uint", &sString, "int", -1, "Uint", &wString, "int", nSize + 1)
        DllCall("ole32\CLSIDFromString", "Uint",&wString , "Uint", &CLSID)
        
        If   hTabUI && DllCall("oleacc\AccessibleObjectFromWindow", "Uint", hTabUI, "Uint",-4, "Uint", &CLSID , "UintP", pacc)=0 
            { 
            pacc := ComObject(9, pacc, 1), ObjAddRef(pacc)
            Loop, %   pacc.accChildCount 
                If   paccChild:=pacc.accChild(A_Index) 
                    If   paccChild.accRole(0+0) = 0x3C 
                        { 
                        paccTab:=paccChild 
                        Break 
                        } 
                    Else   ObjRelease(paccChild) 
            ObjRelease(pacc) 
            } 
        If   pacc:=paccTab 
            { 
            Loop, %   pacc.accChildCount
                If   paccChild:=pacc.accChild(A_Index) 
                    If   paccChild.accName(0+0) = sTitle   
                        { 
                        ObjRelease(pwb)
                        paccChild.accDoDefaultAction(0)
                        ObjRelease(paccChild) 
                        Break 
                        } 
                    Else   ObjRelease(paccChild) 
            ObjRelease(pacc) 
            }  
        WinActivate,% sTitle
        } 


    isInstalled()
        {
        Static IE_path
        
        ;; find where windows believes IE is installed
        ;; certain corp installs may have this in other than expected folders
        if !IE_path
            RegRead, IE_path, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE
        ;~ MsgBox % IE_path
        ;; Perhaps policies prevent reading this key
        if ( ErrorLevel || !IE_path )
            IE_path := "C:\Program Files\Internet Explorer\iexplore.exe"
        
        ;; make sure it installed
        if !FileExist( IE_path )
            {
            MsgBox, 4112, Internet Explorer Not Found, IE does not appear to be installed`nCannot continue `nClick OK to Exit!!!
            ExitApp
            }
        } 

    pDoc()
        {
        return this.IHTMLWindow2_from_IWebDOCUMENT( this.pwb.document ).document
        }

    clean_IE_Title( ByRef sTitle = "" ) 
        {
        return sTitle := RegExReplace( sTitle ? sTitle : this.active_IE_Title(), this.IE_Suffix() "$", "" )
        }

    IE_Suffix() 
        {
        static sIE_Suffix
        if !sIE_Suffix
            {
            ;; HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main
            RegRead, sIE_Suffix, HKCU, Software\Microsoft\Internet Explorer\Main, Window Title ;, Windows Internet Explorer,
            sIE_Suffix := " - " sIE_Suffix
            }
        return sIE_Suffix
        }

    active_IE_Title() ;; returns the title of the topmost browser if exists from the stack
        {
        sTitle := "NO IE Window Open"
        if winexist( "ahk_class IEFrame" )
            {
            titlematchMode := A_TitleMatchMode
            titlematchSpeed := A_TitleMatchModeSpeed
            SetTitleMatchMode, 2	
            SetTitleMatchMode, Slow
            WinGetTitle, sTitle, %sIE_Suffix% ahk_class IEFrame
            SetTitleMatchMode, %titlematchMode%	
            SetTitleMatchMode, %titlematchSpeed%
            }
        return RegExReplace( sTitle, this.IE_Suffix() "$", "" )
        }
        
        
        
    IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT )
        {
        static IID_IHTMLWindow2 := "{332C4427-26CB-11D0-B483-00C04FD90119}"  ; IID_IHTMLWindow2
        return ComObj(9,ComObjQuery( IWebDOCUMENT, IID_IHTMLWindow2, IID_IHTMLWindow2),1)
        }

    IWebDOCUMENT_from_IWebDOCUMENT( IWebDOCUMENT ) ;bypasses certain security issues
        {
        return this.IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT ).document
        }

    IWebBrowserApp_from_IWebDOCUMENT( IWebDOCUMENT )
        {
        static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"  ; IID_IWebBrowserApp
        return ComObj(9,ComObjQuery( this.IHTMLWindow2_from_IWebDOCUMENT( IWebDOCUMENT ), IID_IWebBrowserApp, IID_IWebBrowserApp),1)
        }

    IWebBrowserApp_from_Internet_Explorer_Server_HWND( hwnd, Svr#=1 ) 
        {               ;// based on ComObjQuery docs
        static msg := DllCall( "RegisterWindowMessage", "str", "WM_HTML_GETOBJECT" )
            , IID_IWebDOCUMENT := "{332C4425-26CB-11D0-B483-00C04FD90119}"
        
        SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, ahk_id %hwnd%
        
        if (ErrorLevel != "FAIL") 
            {
            lResult := ErrorLevel
            VarSetCapacity( GUID, 16, 0 )
            if DllCall( "ole32\CLSIDFromString", "wstr", IID_IWebDOCUMENT, "ptr", &GUID ) >= 0 
                {
                DllCall( "oleacc\ObjectFromLresult", "ptr", lResult, "ptr", &GUID, "ptr", 0, "ptr*", IWebDOCUMENT )
                return  this.IWebBrowserApp_from_IWebDOCUMENT( IWebDOCUMENT )
                }
            }
        }
}
