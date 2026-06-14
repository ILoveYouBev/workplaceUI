
prank:=true
;~ https://www.autohotkey.com/boards/viewtopic.php?t=35737  by scriptor2016 ; For voice recognition to work you need Microsoft SAPI installed in your PC, some versions of Windows don't support voice recognition though.
; You may also need to train voice recognition in Windows so that it will understand your voice.
#Persistent
#SingleInstance, Force
#notrayicon

global pspeaker := ComObjCreate("SAPI.SpVoice") ;plistener := ComObjCreate("SAPI.SpSharedRecognizer") 
plistener:= ComObjCreate("SAPI.SpInprocRecognizer") ; For not showing Windows Voice Recognition widget.
paudioinputs := plistener.GetAudioInputs() ; For not showing Windows Voice Recognition widget.
plistener.AudioInput := paudioinputs.Item(0)   ; For not showing Windows Voice Recognition widget.
ObjRelease(paudioinputs) ; Release object from memory, it is not needed anymore.
pcontext := plistener.CreateRecoContext()
pgrammar := pcontext.CreateGrammar()
pgrammar.DictationSetState(0)
prules := pgrammar.Rules()
prulec := prules.Add("wordsRule", 0x1|0x20)
prulec.Clear()
pstate := prulec.InitialState()

;Object of Text responses and their Labels to jump to when detected
global responses:={"Let me see":"TurnOffPrank","What do you mean":"TurnOnPrank"}
for Text, v in Responses ;Need to add each text to the pstate object and watch for them
	pstate.AddWordTransition( ComObjParameter(13,0),Text) ; ComObjParemeter(13,0) is value Null for AHK_L

prules.Commit()
pgrammar.CmdSetRuleState("wordsRule",1)
prules.Commit()
ComObjConnect(pcontext, "On")
If (pspeaker && plistener && pcontext && pgrammar && prules && prulec && pstate){	
	;SplashTextOn,300,50,,Voice recognition initialisation succeeded
}Else { 
	MsgBox, Sorry, voice recognition initialisation FAILED  ;	pspeaker.speak("Starting voice recognition initialisation failed")
}
sleep, 2000
SplashTextOff


return


;********************On recognition function***********************************
OnRecognition(StreamNum,StreamPos,RecogType,Result){
	sText:= Result.PhraseInfo().GetText() ; Grab the text we just spoke and go to that subroutine
		;~ pspeaker.Speak("You said " sText) 	;~ MsgBox, Command is %sText%
	if (Responses[sText]) ;If text is found as a key in the object then... 
		gosub % Responses[sText] ;jump to the gosub
	ObjRelease(sText)
}


return

TurnOnPrank:
Prank:=true
return

TurnOffPrank:
Prank:=False
return

#If prank
!Tab::

return

return
*/

Tab::Reload