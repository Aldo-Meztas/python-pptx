#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 15:54:20 2020

@author: ssccdmx-amh
"""
from pptx import Presentation
# para imagenes se utiliza este modulo
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

class ArchivosPptx():    
    def __init__(self,pathTemplete):
        self.f = open(pathTemplete)
        self.prs =Presentation(pathTemplete)        
        
    def presentacion(self,titulo):
        prs = self.prs
        tituDiapo = prs.slide_layouts[0]       
        slide = prs.slides.add_slide(tituDiapo)        
        title = slide.shapes.title             
        title.text = titulo
        
    def temaDiapo(self,titulo,subtitulo):        
        prs = self.prs
        SLD_LAYOUT_TITLE_AND_CONTENT = 1        
        slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_AND_CONTENT]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title        
        title.text = titulo   
        subtitle = slide.placeholders[1]
        subtitle.text = subtitulo
        
    def diseñoDiapo1(self,titulo,parrafo,pathImg):
        prs = self.prs
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)        
        left = top = width = height = Inches(1)
        
        title = slide.shapes.title        
        title.text = titulo
        
        txBox = slide.shapes.add_textbox(Inches(4.07), Inches(1.93), Inches(1.82), Inches(5.54))        
        tf = txBox.text_frame                
        tf.text = parrafo
        p = tf.add_paragraph()
        img_path = pathImg
        left = Inches(1.00)
        top = Inches(1.67)
        height = Inches(2.45)
        pic = slide.shapes.add_picture(img_path, left, top, height=height)
    
    def diseñoDiapo2(self,titulo,parrafo):
        prs = self.prs
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)        
        
        
        title = slide.shapes.title        
        title.text = titulo
        
        txBox = slide.shapes.add_textbox(Inches(1.00), Inches(1.33), Inches(0.57), Inches(8.30))        
        tf = txBox.text_frame                
        tf.text = parrafo
        p = tf.add_paragraph()
    
    
    def agregarImagen(self,pathImg,descripcion):
        # Ruta de imagen
        img_path = pathImg
        prs = self.prs
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        left = Inches(5)
        height = Inches(3.12)
        
        pic = slide.shapes.add_picture(img_path, Inches(3.14), Inches(0.60), height=height)
        
        # Descripcion de imagen
        txBox = slide.shapes.add_textbox(Inches(1.81), Inches(4.05), Inches(0.81), Inches(5.94))        
        tf = txBox.text_frame                
        tf.text = descripcion
        p = tf.add_paragraph()
    
    def diapoViñetas(self,titulo,tema,subtema,puntos):
        prs = self.prs
        
        bullet_slide_layout = prs.slide_layouts[3]
        
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        
        title_shape.text = titulo
        
        tf = body_shape.text_frame
        tf.text = tema
        
        p = tf.add_paragraph()
        p.text = subtema
        p.level = 1
        
        p = tf.add_paragraph()
        p.text = puntos
        p.level = 1
        
    
        
    def crearTabla(self):
        
        from pptx.util import Inches
        prs = self.prs
        title_only_slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(title_only_slide_layout)
        shapes = slide.shapes
        
        shapes.title.text = 'Añadiendo tablas'
        
        rows = cols = 2
        left = top = Inches(2.0)
        width = Inches(6.0)
        height = Inches(0.8)
        
        table = shapes.add_table(rows, cols, left, top, width, height).table
        
        # set column widths
        table.columns[0].width = Inches(2.0)
        table.columns[1].width = Inches(4.0)
        
        # write column headings
        table.cell(0, 0).text = 'Encabezado-1'
        table.cell(0, 1).text = 'Encabezado-2'
        
        # write body cells
        table.cell(1, 0).text = 'valo-1'
        table.cell(1, 1).text = 'valor-2'
        
    
        
            
        


    def cerrarPlantilla(self):
        self.prs.save('Salida.pptx')    
        self.f.close()

        
# =================EJEMPLO PARA DIAPOSITIVAS=================
miObjeto = ArchivosPptx('templete.pptx')

texto = "Esto es una breve descripcion de la imagen por texto \n"+"hola Mundo"
img = 'descarga.png'

miObjeto.presentacion('presentacion')
miObjeto.temaDiapo("Diapositiva", "subtitulos")
miObjeto.diseñoDiapo1('titulo', 'parrafo', img)
miObjeto.agregarImagen(img,texto)
miObjeto.diseñoDiapo2('titulo', texto)

miObjeto.cerrarPlantilla()
# ===================FIN EJEMPLO===================