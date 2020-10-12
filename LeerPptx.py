#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 12 15:54:20 2020

@author: ssccdmx-amh
"""
from pptx import Presentation
# para imagenegenes se utiliza este modulo
from pptx.util import Inches

class ArchivosPptx():    
    def primerFormato(self):
        prs = Presentation()
        tituDiapo = prs.slide_layouts[0]       
        slide = prs.slides.add_slide(tituDiapo)        
        title = slide.shapes.title
        subtitle = slide.placeholders[1]        
        title.text = "Hola Mundo!"
        subtitle.text = "Subtitulos!"        
        prs.save('Formato1.pptx')    
        
    def segundoFormato(self):
        prs = Presentation()
        bullet_slide_layout = prs.slide_layouts[1]
        
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        
        title_shape.text = 'Diapositiva de viñeta'
        
        tf = body_shape.text_frame
        tf.text = 'Find the bullet slide layout'
        
        p = tf.add_paragraph()
        p.text = 'Use _TextFrame.text for first bullet'
        p.level = 1
        
        p = tf.add_paragraph()
        p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
        p.level = 2

        prs.save('Formato2.pptx')
    def agregarImagen(self):
        # Ruta de imagen
        img_path = 'descarga.png'

        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        left = top = Inches(1)
        pic = slide.shapes.add_picture(img_path, left, top)
        
        left = Inches(5)
        height = Inches(5.5)
        pic = slide.shapes.add_picture(img_path, left, top, height=height)

        prs.save('creacionImagen.pptx')
        
    def crearTabla(self):
        
        from pptx.util import Inches
        prs = Presentation()
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
        
        prs.save('creacionTabla.pptx')
        

miObjeto = ArchivosPptx()
miObjeto.primerFormato()
miObjeto.segundoFormato()
miObjeto.agregarImagen()
miObjeto.crearTabla()