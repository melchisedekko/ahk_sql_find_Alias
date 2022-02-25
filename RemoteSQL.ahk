Class RemoteSQL{
	__New(ConnectString,Timeout:=1500){
		OnExit,ExitRSQL
		this.Connection:=ComObjCreate("ADODB.Connection")
		this.ConnectString:=ConnectString,this.Connect()
		this.CloseTimer:=this.Close.Bind(this)
		this.Timeout:=-Abs(TimeOut)
		this.ConnectionState:=0
	}CheckConnection(){
		if(!this.Connection.State)
			this.Connection.Open()
	}Clean(Text:="",Wrap:="'",Empty:=0){
		static WrapChar:={(Chr(34)):{Wrap:Chr(34),Regex:"(\x22)"},"[":"]","'":{Regex:"(')",Wrap:"'"},"[":{Regex:"(\[|\])",Wrap:"]"}}
		if(Text=""||Text="NULL")
			return (Empty?"''":"NULL")
		Text:=RegExReplace(Text,WrapChar[Wrap].Regex,"$1$1")
		return Wrap Text WrapChar[Wrap].Wrap
	}Close(){
		if(this.Connection.State)
		this.Connection.Close()
	}Connect(){
		this.Connection.ConnectionString:=this.ConnectString
		this.Connection.Open()
		for a,b in this.Connection.Errors
			if (a.Number != 0) {
			m(a.Number,a.Description,a.Source)
			}
		if(this.Connection.State!=1){
			ErrorLevel := "-1"
			this.Connection:=""
			return,ErrorLevel
		}this.Close()
	}Exec(Query){
		return this.Query(Query)
	}Exit(){
		ExitRSQL:
		this.Close()
		ExitApp
		return
	}Query(SQL){
		ComObjError(0),MySQLClose:=this.CloseTimer
		SetTimer,%MySQLClose%,Off
		Sleep,10
		Obj:=[],this.CheckConnection(),RS:=this.Connection.Execute(SQL) ;Execute SQL and return Record Set Object
		if(RS.BOF=-1&&RS.EOF=-1)
			return {Empty:"No Results"},this.Close()
		for a in this.Connection.Errors{
			Try, DebugWindow("the query wich resulted in error was => " . SQL,,1)
			m("Function: " A_ThisFunc,"Line: " A_LineNumber,"","a.Description: " a.Description,"a.HelpContent: " a.HelpContent,"a.HelpFile: " a.HelpFile,"a.NativeError: " a.NativeError,"a.Number: " a.Number,"a.Source: " a.Source,"a.SQLState: " a.SQLState,"",SQL)
			return this.Close()
		}
		if(RS.BOF&&RS.EOF){
			m("Function: " A_ThisFunc,"Line: " A_LineNumber,"this.Connection.Errors",this.Connection.Errors.Count(),"RS.Errors: " RS.Errors.HelpFile,"RS.BOF: " RS.BOF,"RS.EOF: " RS.EOF,"RS.Fields.Count: " RS.Fields.Count())
			for a,b in RS.Errors
				m("Function: " A_ThisFunc,"Line: " A_LineNumber,"",a.Description)
			return m("Function: " A_ThisFunc,"Line: " A_LineNumber,"","HERE!",SQL),this.Close()
		}
		if(RS.BOF&&RS.EOF){
			ErrorLevel := "-2"
			return ErrorLevel,this.Close()
		}if(!RS.Fields.Count){
			return this.Close()
		}
		while(!RS.EOF){
			RSFields:=RS.Fields
			Index:=A_Index
			Loop,% RSFields.Count{
				RSField := RSFields.Item(A_Index-1)
				if(RSField.Type=205){
					;~ m(RSField.Size)
					Stream:=ComObjCreate("ADODB.Stream")
					Stream.Type:=1
					Stream.Open()
					Stream.Write(RSField.Value)
					;~ m("Function: " A_ThisFunc,"Line: " A_LineNumber,"","HERE!!!!")
					Stream.SaveToFile(Images "Image "(A_TickCount)".jpg")
					Stream.Close()
				}else{
					Obj[Index,RSField.Name]:=Trim(RSfield.Value)
				}
			}
			RS.MoveNext()
		}
		SetTimer,%MySQLClose%,% this.TimeOut
		ComObjError(1)
		return Obj
	}
}