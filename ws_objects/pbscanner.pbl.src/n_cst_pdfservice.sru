$PBExportHeader$n_cst_pdfservice.sru
forward
global type n_cst_pdfservice from nonvisualobject
end type
end forward

global type n_cst_pdfservice from nonvisualobject
end type
global n_cst_pdfservice n_cst_pdfservice

forward prototypes
public function string of_cambiar_extension (string as_source, string as_newformat)
public function string of_imagetopdf (string as_image)
public function string of_imagearraytopdf (string as_image[])
public subroutine of_pdfimage (pdfdocument a_pdf_doc, string as_image)
end prototypes

public function string of_cambiar_extension (string as_source, string as_newformat);String ls_oldFormat
Long ll_rtn
Integer li_FormatLen
String ls_newFileName

li_FormatLen = Len(as_source) - LastPos(as_source, ".") + 1

ls_oldFormat = lower(mid(as_source, LastPos(as_source, "."),  li_FormatLen))

ls_newFileName = replace(as_source, pos(as_source, ls_oldFormat), li_FormatLen, as_newformat)

RETURN ls_newFileName
end function

public function string of_imagetopdf (string as_image);String ls_pdf
PDFdocument lpdf_doc
Long ll_rtn

ls_pdf=of_cambiar_extension(as_image, ".pdf")
	
IF FileExists(ls_pdf) THEN FileDelete(ls_pdf)
	
lpdf_doc = Create PDFdocument
	
//Creamos la Imagen en el Objeto PdfImage
of_pdfimage(lpdf_doc, as_image)

ll_rtn = lpdf_doc.save(ls_pdf)

IF ll_rtn <> 1 THEN
	Messagebox("Error", "Error saving pdf", Exclamation!)
	RETURN ""
END IF	

IF NOT FileDelete(as_image) THEN
	Messagebox("Error", "Error deleting image", Exclamation!)
	RETURN ""
END IF	
	
DESTROY lpdf_doc

RETURN ls_pdf
end function

public function string of_imagearraytopdf (string as_image[]);String ls_pdf
PDFdocument lpdf_doc
Long ll_rtn
Integer li_image, li_TotalImages

li_TotalImages = UpperBound(as_image[])

ls_pdf = of_cambiar_extension(as_image[1], ".pdf")

ls_pdf = replace(ls_pdf, pos(ls_pdf, "_1"), 2, "")

IF FileExists(ls_pdf) THEN FileDelete(ls_pdf)
	
lpdf_doc = Create PDFdocument

For li_image = 1 to li_TotalImages
	of_pdfimage( lpdf_doc, as_image[li_image])
Next	
	
ll_rtn = lpdf_doc.save(ls_pdf)

IF ll_rtn <> 1 THEN
	Messagebox("Error", "Error saving pdf", Exclamation!)
	RETURN ""
END IF	

For li_image = 1 to li_TotalImages
	IF NOT FileDelete(as_image[li_image]) THEN
		Messagebox("Error", "Error deleting image", Exclamation!)
		RETURN ""
	END IF	
Next		

DESTROY lpdf_doc
RETURN ls_pdf
end function

public subroutine of_pdfimage (pdfdocument a_pdf_doc, string as_image);String ls_pdf
PDFpage lpdf_page
PDFImage lpdf_image
Long ll_rtn

lpdf_page = Create PDFpage
lpdf_image = Create PDFImage
	
//Creamos la Imagen en el Objeto PdfImage
lpdf_image.filename = as_image
lpdf_image.x=0
lpdf_image.y=0
lpdf_image.height=842 //A4 height
lpdf_image.width=595 //A4 Width
//lpdf_image.FitMethod = PDFImageFitmethod_Clip!  //Ajustar la imagen con recorte.
//lpdf_image.FitMethod = PDFImageFitmethod_Entire!  //Ajustar la imagen sin recorte.
lpdf_image.FitMethod = PDFImageFitmethod_Meet!  //Ajustar la imagen con cambio de tamaño proporcional.
				
//Añadimos la Imagen Creada al Objeto PdfPage
ll_rtn = lpdf_page.addcontent( lpdf_image)
	
IF ll_rtn <> 1 THEN
	Messagebox("Error", "Error inserting image", Exclamation!)
	RETURN 
END IF	
		
//Importamos la Pagina con la Imagen al Objeto PDFDocument
ll_rtn = a_pdf_doc.AddPage(lpdf_page)
	
IF ll_rtn <> 1 THEN
	Messagebox("Error", "Error inserting page", Exclamation!)
	RETURN 
END IF	

a_pdf_doc.properties.application = "RSRSYSTEM PbScanner2022"
a_pdf_doc.properties.author = "Ramón San Félix Ramón"
a_pdf_doc.properties.keywords = "Escaner, PDF, Imagenes, WIA"
a_pdf_doc.properties.subject = "https://rsrsystem.blogspot.com/2023/12/escanear-en-powerbuilder-2022r2-wia-c.html"
a_pdf_doc.properties.title = "Escanear en PowerBuilder 2022R3 (WIA C#)"
	
DESTROY  lpdf_page 
DESTROY lpdf_image

end subroutine

on n_cst_pdfservice.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_pdfservice.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

