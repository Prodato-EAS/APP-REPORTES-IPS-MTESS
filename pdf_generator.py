"""
Fast PDF generation using ReportLab
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from datetime import datetime, timezone, timedelta
from io import BytesIO
import urllib.request


def generate_pdf_reportlab(company_data, inconsistencies, company_ruc, company_name):
    """
    Generate PDF using ReportLab with exact same design as WeasyPrint version
    
    Args:
        company_data: DataFrame with all company records
        inconsistencies: DataFrame with inconsistent records
        company_ruc: Company RUC number
        company_name: Company name
    
    Returns:
        BytesIO buffer containing the PDF
    """
    
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        page_num = canvas.getPageNumber()
        text = "Página %d" % page_num
        canvas.drawString(1*cm, 0.75*cm, text)
        canvas.restoreState()
    
    # Generate timestamp (Paraguay time UTC-3)
    py_tz = timezone(timedelta(hours=-3))
    now = datetime.now(py_tz)
    timestamp = now.strftime("%d-%m-%Y %H:%M")
    
    total_records = len(company_data)
    total_inconsistencies = len(inconsistencies)
    inconsistencies_text = f"{total_inconsistencies}" if total_inconsistencies > 0 else "No se encontraron inconsistencias"
    
    # Create PDF buffer
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, 
                           leftMargin=1*cm, rightMargin=1*cm, title=f"Reporte - {company_name}")
    
    # Build content
    story = []
    
    # Define colors
    color_gray_text = colors.HexColor('#3a3a3a')
    color_gray_label = colors.HexColor('#3a3a3a')  
    color_gray_title = colors.HexColor('#5f6368')
    color_border = colors.HexColor('#e5e5e5')
    color_bg_summary = colors.HexColor('#f9fafb')
    color_bg_header = colors.HexColor('#f8f9fa')
    color_badge_success_bg = colors.HexColor('#e6f4ea')
    color_badge_success_text = colors.HexColor('#137333')
    color_badge_danger_bg = colors.HexColor('#fce8e6')
    color_badge_danger_text = colors.HexColor('#c5221f')
    color_badge_info_bg = colors.HexColor('#e8f0fe')
    color_badge_info_text = colors.HexColor('#1967d2')
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Logo
    try:
        logo_url = "https://prodato.com.py/wp-content/uploads/2022/06/cropped-logo.jpg"
        logo_data = BytesIO(urllib.request.urlopen(logo_url, timeout=3).read())
        logo = Image(logo_data, width=50*mm, height=15*mm, kind='proportional')
        logo.hAlign = 'LEFT'
        story.append(logo)
        story.append(Spacer(1, 5))
    except:
        pass
    
    # Title
    title_style = ParagraphStyle('Title', parent=styles['Heading3'], fontSize=10, 
                                 textColor=color_gray_title, fontName='Helvetica-Bold')
    story.append(Paragraph("Datos de la empresa", title_style))
    story.append(Spacer(1, 5))
    
    # Company info grid
    info_style = ParagraphStyle('Info', fontSize=8, textColor=color_gray_text, fontName='Helvetica')
    label_style = ParagraphStyle('Label', fontSize=8, textColor=color_gray_label, fontName='Helvetica-Bold')
    
    info_data = [
        [Paragraph("RUC:", label_style), Paragraph(str(company_ruc), info_style)],
        [Paragraph("Empresa:", label_style), Paragraph(company_name, info_style)],
        [Paragraph("Fecha reporte:", label_style), Paragraph(timestamp, info_style)]
    ]
    info_table = Table(info_data, colWidths=[2*cm, 16*cm])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(info_table)
    
    # Border line
    story.append(Spacer(1, 3))
    line_table = Table([['']], colWidths=[19*cm])
    line_table.setStyle(TableStyle([
        ('LINEBELOW', (0, 0), (-1, -1), 1, color_border),
    ]))
    story.append(line_table)
    story.append(Spacer(1, 12))
    
    # Summary box
    summary_style = ParagraphStyle('Summary', fontSize=9, textColor=color_gray_title, 
                                  fontName='Helvetica-Bold')
    summary_item_style = ParagraphStyle('SummaryItem', fontSize=8, textColor=colors.HexColor('#70757a'), 
                                       fontName='Helvetica')
    
    summary_data = [
        [Paragraph("Resumen", summary_style)],
        [Paragraph(f"Datos generales: {total_records} registro(s)", summary_item_style)],
        [Paragraph(f"Inconsistencias detectadas: {inconsistencies_text}", summary_item_style)]
    ]
    summary_table = Table(summary_data, colWidths=[19*cm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), color_bg_summary),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e8eaed')),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 15))
    
    # Section title
    section_style = ParagraphStyle('Section', fontSize=9, textColor=color_gray_title, 
                                  fontName='Helvetica-Bold')
    story.append(Paragraph("DATOS GENERALES", section_style))
    story.append(Spacer(1, 10))
    
    # Helper function for badges
    def format_badge_text(val):
        val_upper = str(val).upper()
        if val_upper == 'COINCIDE':
            return ('Coincide', color_badge_success_bg, color_badge_success_text)
        elif val_upper == 'NO_COINCIDE':
            return ('No Coincide', color_badge_danger_bg, color_badge_danger_text)
        elif 'SIN_REGISTRO' in val_upper:
            return (str(val), color_badge_danger_bg, color_badge_danger_text)
        elif val_upper == 'VIGENTE':
            return ('Vigente', color_badge_info_bg, color_badge_info_text)
        else:
            return (str(val) if val else '-', None, color_gray_text)
    
    # Data table styles
    cell_style = ParagraphStyle('Cell', fontSize=6.2, textColor=color_gray_text, 
                               fontName='Helvetica', leading=7)
    header_style = ParagraphStyle('Header', fontSize=6, textColor=color_gray_title, 
                                 fontName='Helvetica-Bold', leading=7)
    badge_style = ParagraphStyle('Badge', fontSize=5.5, fontName='Helvetica-Bold', leading=6.5)
    
    # Column widths (percentages converted to cm, total ~19cm)
    col_widths = [0.57*cm, 1.52*cm, 2.66*cm, 1.33*cm, 1.33*cm, 1.71*cm, 
                 1.33*cm, 1.33*cm, 1.71*cm, 1.33*cm, 1.33*cm, 1.71*cm]
    
    # Headers
    headers = ['N°', 'Cédula', 'Nombre', 'Est. IPS', 'Est. MTESS', 'Aud. Est.',
              'Ent. IPS', 'Ent. MTESS', 'Aud. Ent.', 'Sal. IPS', 'Sal. MTESS', 'Aud. Sal.']
    table_data = [[Paragraph(h, header_style) for h in headers]]
    
    # Helper to create badge paragraphs
    def make_badge(badge_info):
        text, bg, fg = badge_info
        if bg:
            # Use alignment and padding to prevent background overflow
            return Paragraph(f'<para backColor="{bg.hexval()}" textColor="{fg.hexval()}" fontSize="5.5" fontName="Helvetica-Bold" alignment="center">{text}</para>', badge_style)
        else:
            return Paragraph(text, cell_style)
    
    # Data rows
    for idx, row in enumerate(company_data.itertuples(), 1):
        # Format audit badges
        aud_estado = format_badge_text(getattr(row, 'AUD_ESTADO', ''))
        aud_entrada = format_badge_text(getattr(row, 'AUD_ENTRADA', ''))
        aud_salida = format_badge_text(getattr(row, 'AUD_SALIDA', ''))
        
        row_data = [
            Paragraph(str(idx), cell_style),
            Paragraph(str(getattr(row, 'Cedula', '')), cell_style),
            Paragraph(str(getattr(row, 'Nombre', '')), cell_style),
            Paragraph(str(getattr(row, 'Estado_IPS', '-')), cell_style),
            Paragraph(str(getattr(row, 'Estado_MTESS', '-')), cell_style),
            make_badge(aud_estado),
            Paragraph(str(getattr(row, 'Entrada_IPS', '-')), cell_style),
            Paragraph(str(getattr(row, 'Entrada_MTESS', '-')), cell_style),
            make_badge(aud_entrada),
            Paragraph(str(getattr(row, 'Salida_IPS', '-')), cell_style),
            Paragraph(str(getattr(row, 'Salida_MTESS', '-')), cell_style),
            make_badge(aud_salida)
        ]
        table_data.append(row_data)
    
    # Create table
    data_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    data_table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), color_bg_header),
        ('TEXTCOLOR', (0, 0), (-1, 0), color_gray_title),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 6),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor('#e8eaed')),
        # Cell styling
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 6.2),
        ('TEXTCOLOR', (0, 1), (-1, -1), color_gray_text),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, colors.HexColor('#f5f5f5')),
        # Alternating rows
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#fafafa')]),
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    story.append(data_table)
    story.append(Spacer(1, 20))
    
    # Inconsistencies section
    story.append(PageBreak())
    story.append(Paragraph("INCONSISTENCIAS DETECTADAS", section_style))
    story.append(Spacer(1, 10))
    
    if not inconsistencies.empty:
        # Build inconsistencies table
        incon_table_data = [[Paragraph(h, header_style) for h in headers]]
        
        for idx, row in enumerate(inconsistencies.itertuples(), 1):
            aud_estado = format_badge_text(getattr(row, 'AUD_ESTADO', ''))
            aud_entrada = format_badge_text(getattr(row, 'AUD_ENTRADA', ''))
            aud_salida = format_badge_text(getattr(row, 'AUD_SALIDA', ''))
            
            row_data = [
                Paragraph(str(idx), cell_style),
                Paragraph(str(getattr(row, 'Cedula', '')), cell_style),
                Paragraph(str(getattr(row, 'Nombre', '')), cell_style),
                Paragraph(str(getattr(row, 'Estado_IPS', '-')), cell_style),
                Paragraph(str(getattr(row, 'Estado_MTESS', '-')), cell_style),
                make_badge(aud_estado),
                Paragraph(str(getattr(row, 'Entrada_IPS', '-')), cell_style),
                Paragraph(str(getattr(row, 'Entrada_MTESS', '-')), cell_style),
                make_badge(aud_entrada),
                Paragraph(str(getattr(row, 'Salida_IPS', '-')), cell_style),
                Paragraph(str(getattr(row, 'Salida_MTESS', '-')), cell_style),
                make_badge(aud_salida)
            ]
            incon_table_data.append(row_data)
        
        incon_table = Table(incon_table_data, colWidths=col_widths, repeatRows=1)
        incon_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), color_bg_header),
            ('TEXTCOLOR', (0, 0), (-1, 0), color_gray_title),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 6),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor('#e8eaed')),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6.2),
            ('TEXTCOLOR', (0, 1), (-1, -1), color_gray_text),
            ('LINEBELOW', (0, 1), (-1, -1), 0.5, colors.HexColor('#f5f5f5')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#fafafa')]),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        story.append(incon_table)
    else:
        # No inconsistencies message
        no_incon_style = ParagraphStyle('NoIncon', fontSize=8, textColor=colors.HexColor('#155724'), 
                                       fontName='Helvetica')
        no_incon_data = [[Paragraph("<b>Todo en orden:</b> No se detectaron inconsistencias en esta empresa.", no_incon_style)]]
        no_incon_table = Table(no_incon_data, colWidths=[19*cm])
        no_incon_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#d4edda')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#c3e6cb')),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(no_incon_table)
    
    # Build PDF
    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    
    # Return buffer
    buffer.seek(0)
    return buffer