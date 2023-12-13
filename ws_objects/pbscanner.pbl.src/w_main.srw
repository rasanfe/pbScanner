$PBExportHeader$w_main.srw
forward
global type w_main from window
end type
type wb_image from webbrowser within w_main
end type
type st_output from statictext within w_main
end type
type sle_output from singlelineedit within w_main
end type
type ddlb_format from dropdownlistbox within w_main
end type
type st_format from statictext within w_main
end type
type st_2 from statictext within w_main
end type
type lb_devices from listbox within w_main
end type
type cb_scan from commandbutton within w_main
end type
type st_filename from statictext within w_main
end type
type sle_filename from singlelineedit within w_main
end type
type p_2 from picture within w_main
end type
type cb_change from commandbutton within w_main
end type
type st_info from statictext within w_main
end type
type st_myversion from statictext within w_main
end type
type st_platform from statictext within w_main
end type
type gb_1 from groupbox within w_main
end type
type r_2 from rectangle within w_main
end type
end forward

global type w_main from window
integer width = 4398
integer height = 3156
boolean titlebar = true
string title = "WIA Scanner Example"
boolean controlmenu = true
boolean minbox = true
string icon = "AppIcon!"
boolean center = true
wb_image wb_image
st_output st_output
sle_output sle_output
ddlb_format ddlb_format
st_format st_format
st_2 st_2
lb_devices lb_devices
cb_scan cb_scan
st_filename st_filename
sle_filename sle_filename
p_2 p_2
cb_change cb_change
st_info st_info
st_myversion st_myversion
st_platform st_platform
gb_1 gb_1
r_2 r_2
end type
global w_main w_main

type prototypes
Function long GetFolderDialog(long hwndOwner, string wsDlgTitle, string wsStartPath, ref string wsFolder) Library "getfolderdialog_x64.dll" alias for "GetFolderDialog"
end prototypes

type variables
nvo_scannerwia in_wia
end variables

forward prototypes
public subroutine wf_version (statictext ast_version, statictext ast_patform)
public function integer wf_getfolder (string as_dlgtitle, ref string as_folder)
end prototypes

public subroutine wf_version (statictext ast_version, statictext ast_patform);String ls_version, ls_platform
environment env
integer rtn

rtn = GetEnvironment(env)

IF rtn <> 1 THEN 
	ls_version = string(year(today()))
	ls_platform="32"
ELSE
	ls_version = "20"+ string(env.pbmajorrevision)+ "." + string(env.pbbuildnumber)
	ls_platform=string(env.ProcessBitness)
END IF

ls_platform += " Bits"

ast_version.text=ls_version
ast_patform.text=ls_platform

end subroutine

public function integer wf_getfolder (string as_dlgtitle, ref string as_folder);/*
Copyright 2020 Davros

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

Constant long MAX_PATH = 260
Constant long GFD_OK = 0
Constant long GFD_ERROR = -1
long ll_hwndOwner, ll_ret
string ls_startPath

ll_hwndOwner = Handle(This)

If as_folder = "" Then
	ls_startPath =gs_dir
Else
	ls_startPath =as_folder
End IF	

If Not DirectoryExists(ls_startPath) Then
	ls_startPath = Space(MAX_PATH)
End If

as_folder = Space(MAX_PATH)

ll_ret = GetFolderDialog(ll_hwndOwner, as_dlgTitle, ls_startPath, REF as_folder)

IF ll_ret=0 Then ll_ret = 1

Return ll_ret
end function

on w_main.create
this.wb_image=create wb_image
this.st_output=create st_output
this.sle_output=create sle_output
this.ddlb_format=create ddlb_format
this.st_format=create st_format
this.st_2=create st_2
this.lb_devices=create lb_devices
this.cb_scan=create cb_scan
this.st_filename=create st_filename
this.sle_filename=create sle_filename
this.p_2=create p_2
this.cb_change=create cb_change
this.st_info=create st_info
this.st_myversion=create st_myversion
this.st_platform=create st_platform
this.gb_1=create gb_1
this.r_2=create r_2
this.Control[]={this.wb_image,&
this.st_output,&
this.sle_output,&
this.ddlb_format,&
this.st_format,&
this.st_2,&
this.lb_devices,&
this.cb_scan,&
this.st_filename,&
this.sle_filename,&
this.p_2,&
this.cb_change,&
this.st_info,&
this.st_myversion,&
this.st_platform,&
this.gb_1,&
this.r_2}
end on

on w_main.destroy
destroy(this.wb_image)
destroy(this.st_output)
destroy(this.sle_output)
destroy(this.ddlb_format)
destroy(this.st_format)
destroy(this.st_2)
destroy(this.lb_devices)
destroy(this.cb_scan)
destroy(this.st_filename)
destroy(this.sle_filename)
destroy(this.p_2)
destroy(this.cb_change)
destroy(this.st_info)
destroy(this.st_myversion)
destroy(this.st_platform)
destroy(this.gb_1)
destroy(this.r_2)
end on

event open;String ls_devices[]
Integer li_device, li_devices

wf_version(st_myversion, st_platform)



in_wia =  CREATE  nvo_scannerwia


ls_devices[] =  in_wia.of_ListScanners() 

li_devices = UpperBound(ls_devices[])

For li_device = 1 To li_devices
	//IF lb_devices.FindItem (ls_devices[li_device], 0) > 0 THEN CONTINUE
	lb_devices.AddItem(ls_devices[li_device])
Next 	
			
lb_devices.SelectItem(1)
sle_output.text=gs_dir
ddlb_format.SelectItem(ddlb_format.FindItem( "JPEG", 1))
end event

event closequery;Destroy in_wia
end event

type wb_image from webbrowser within w_main
integer x = 1806
integer y = 532
integer width = 2546
integer height = 2424
end type

type st_output from statictext within w_main
integer x = 114
integer y = 1948
integer width = 535
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "Ouput scanned images"
boolean focusrectangle = false
end type

type sle_output from singlelineedit within w_main
integer x = 105
integer y = 2024
integer width = 1646
integer height = 92
integer taborder = 50
integer textsize = -8
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
borderstyle borderstyle = stylelowered!
end type

type ddlb_format from dropdownlistbox within w_main
integer x = 105
integer y = 1500
integer width = 1646
integer height = 420
integer taborder = 50
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
boolean sorted = false
string item[] = {"PNG","JPEG","BMP","GIF","TIFF","PDF"}
borderstyle borderstyle = stylelowered!
end type

type st_format from statictext within w_main
integer x = 114
integer y = 1432
integer width = 329
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "Image format"
boolean focusrectangle = false
end type

type st_2 from statictext within w_main
integer x = 114
integer y = 352
integer width = 370
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "Select a scanner"
boolean focusrectangle = false
end type

type lb_devices from listbox within w_main
integer x = 105
integer y = 432
integer width = 1646
integer height = 616
integer taborder = 40
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
borderstyle borderstyle = stylelowered!
end type

type cb_scan from commandbutton within w_main
integer x = 1806
integer y = 300
integer width = 2546
integer height = 212
integer taborder = 30
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
string text = "Start scan"
end type

event clicked;String ls_filename, ls_format
String ls_Scan, ls_device, ls_outputPath

ls_device = lb_devices.SelectedItem()
ls_format = ddlb_format.text
ls_outputPath = sle_output.text
ls_filename = sle_filename.text

IF ls_format = "PDF" THEN
	ls_format = "PNG"
	ls_Scan = in_wia.of_Scan(ls_device, ls_format, ls_outputPath, ls_filename) 
	ls_Scan = gf_ImageToPdf(ls_scan)
ELSE
	ls_Scan =  in_wia.of_Scan(ls_device, ls_format, ls_outputPath, ls_filename) 
END IF	
				
//Checks the result
If in_wia.il_ErrorType < 0 Then
  messagebox("Failed", in_wia.is_ErrorText)
 wb_image.Navigate("")
  Return
End If

wb_image.Navigate(ls_Scan)


end event

type st_filename from statictext within w_main
integer x = 114
integer y = 1128
integer width = 329
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "Filename"
boolean focusrectangle = false
end type

type sle_filename from singlelineedit within w_main
integer x = 105
integer y = 1204
integer width = 1646
integer height = 92
integer taborder = 40
integer textsize = -8
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
string text = "myscan"
borderstyle borderstyle = stylelowered!
end type

type p_2 from picture within w_main
integer x = 5
integer y = 4
integer width = 1253
integer height = 248
boolean originalsize = true
string picturename = "logo.jpg"
boolean focusrectangle = false
end type

type cb_change from commandbutton within w_main
integer x = 105
integer y = 2140
integer width = 1646
integer height = 116
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Change output folder"
end type

event clicked;integer li_rtn
string ls_path

ls_path= sle_output.text

li_rtn=wf_getfolder("Select Destination Directory", REF ls_path)
	
If li_rtn=1 Then
	sle_output.text=ls_path
End If


end event

type st_info from statictext within w_main
integer x = 3072
integer y = 3000
integer width = 1289
integer height = 52
integer textsize = -7
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 8421504
long backcolor = 553648127
string text = "Copyright © Ramón San Félix Ramón  rsrsystem.soft@gmail.com"
boolean focusrectangle = false
end type

type st_myversion from statictext within w_main
integer x = 3817
integer y = 44
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Versión"
alignment alignment = right!
boolean focusrectangle = false
end type

type st_platform from statictext within w_main
integer x = 3817
integer y = 132
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Bits"
alignment alignment = right!
boolean focusrectangle = false
end type

type gb_1 from groupbox within w_main
integer x = 32
integer y = 276
integer width = 1751
integer height = 2680
integer taborder = 40
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 33554432
string text = "Propieties"
end type

type r_2 from rectangle within w_main
long linecolor = 33554432
linestyle linestyle = transparent!
integer linethickness = 4
long fillcolor = 33521664
integer width = 4402
integer height = 260
end type

