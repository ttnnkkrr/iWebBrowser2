oIE := new iWebBrowser2

class iWebBrowser2 {

    oIE_new(ByRef sTitle = "", ByRef iHWND = "", ByRef sURL = "about:blank", ByRef sHTML = "", bVisible = true)
        {
        this.isInstalled()
        oIE := ComObjCreate("internetexplorer.application")
        oIE.Visible := bVisible
        if sURL
            oIE.Navigate(sURL)
        return oIE
        }

    oIE_get( ByRef sTitle = "", ByRef iHWND = "", ByRef sURL = "", ByRef sHTML = "" )
        {
        this.isInstalled()
        ;~ this function is pointless if no instance of IE is open
        ;~ one edit you might make is to have this function open IE and maybe go to the home page
        if ( !winexist( "ahk_class IEFrame" ) )
            {
            MsgBox, 4112, NO IE Window Found, The Macro will end
            ExitApp
            }
        
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
                if !window.document
                    continue
                pdoc := this.pDoc(window)
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
            ExitApp
            }
            
        return oIE
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

    pDoc(oIE)
        {
        return this.IHTMLWindow2_from_IWebDOCUMENT( oIE.document ).document
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
