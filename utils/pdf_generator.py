from datetime import datetime
from io import BytesIO
from itertools import zip_longest

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus.flowables import HRFlowable

def generate_summary_report_pdf(
    usa_review_data,
    uk_review_data,
    usa_brands,
    uk_brands,
    usa_platforms,
    uk_platforms,
    printing_stats,
    copyright_stats,
    a_plus,
    selected_month=None,
    start_year=None,
    end_year=None,
    filename=None
):
    """
    Generate a PDF summary report with proper year / range handling
    """

    if selected_month and start_year and end_year:
        title_text = f"{selected_month} ({start_year}–{end_year}) Summary Report"
        filename = f"{selected_month}_{start_year}_{end_year}_Summary_Report.pdf"

    elif selected_month and start_year:
        title_text = f"{selected_month} {start_year} Summary Report"
        filename = f"{selected_month}_{start_year}_Summary_Report.pdf"

    elif start_year and end_year:
        title_text = f"{start_year}–{end_year} Summary Report"
        filename = f"{start_year}_{end_year}_Summary_Report.pdf"

    elif start_year:
        title_text = f"{start_year} Summary Report"
        filename = f"{start_year}_Summary_Report.pdf"

    else:
        title_text = "Summary Report"
        filename = "Summary_Report.pdf"

    filename = filename.replace(" ", "_")

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=18
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )

    section_style = ParagraphStyle(
        'Section',
        parent=styles['Heading2'],
        fontSize=16,
        spaceBefore=20,
        spaceAfter=12,
        textColor=colors.darkgreen
    )

    subsection_style = ParagraphStyle(
        'SubSection',
        parent=styles['Heading3'],
        fontSize=12,
        spaceBefore=12,
        spaceAfter=8,
        textColor=colors.darkblue
    )

    story = []
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 20))

    usa_total = sum(usa_review_data.values())
    uk_total = sum(uk_review_data.values())

    usa_attained = usa_review_data.get("Attained", 0)
    uk_attained = uk_review_data.get("Attained", 0)

    combined_total = usa_total + uk_total
    combined_attained = usa_attained + uk_attained

    usa_pct = (usa_attained / usa_total * 100) if usa_total else 0
    uk_pct = (uk_attained / uk_total * 100) if uk_total else 0
    combined_pct = (combined_attained / combined_total * 100) if combined_total else 0

    story.append(Paragraph("📝 Review Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))
    story.append(Spacer(1, 12))

    review_table = Table([
        ["Region", "Total Reviews", "Attained", "Success Rate"],
        ["USA", f"{usa_total:,}", f"{usa_attained:,}", f"{usa_pct:.1f}%"],
        ["UK", f"{uk_total:,}", f"{uk_attained:,}", f"{uk_pct:.1f}%"],
        ["Combined", f"{combined_total:,}", f"{combined_attained:,}", f"{combined_pct:.1f}%"],
    ])

    review_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
    ]))

    story.append(review_table)

    story.append(Spacer(1, 20))
    story.append(Paragraph("📱 Platform Distribution", subsection_style))

    for label, platforms in [("USA", usa_platforms), ("UK", uk_platforms)]:
        story.append(Paragraph(f"{label} Platforms", styles['Normal']))
        table = Table([["Platform", "Count"]] + [[k, v] for k, v in platforms.items()])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

    story.append(Paragraph("🏷️ Brand Performance", subsection_style))

    brand_table = Table(
        [["USA Brand", "Count", "UK Brand", "Count"]] +
        list(zip_longest(
            list(usa_brands.keys()) + ["Total"],
            list(usa_brands.values()) + [sum(usa_brands.values())],
            list(uk_brands.keys()) + ["Total"],
            list(uk_brands.values()) + [sum(uk_brands.values())]
        ))
    )

    brand_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
    ]))

    story.append(brand_table)

    story.append(PageBreak())
    story.append(Paragraph("🖨️ Printing Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))

    printing_table = Table([
        ["Metric", "Value"],
        ["Total Copies", f"{printing_stats['Total_copies']:,}"],
        ["Total Cost", f"${printing_stats['Total_cost']:,.2f}"],
        ["Avg Cost/Copy", f"${printing_stats['Average']:.2f}"]
    ])

    printing_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    story.append(printing_table)
    story.append(Spacer(1, 20))
    story.append(Paragraph("©️ Copyright Analytics", section_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))

    success_rate = (
        copyright_stats['result_count'] /
        copyright_stats['Total_copyrights'] * 100
        if copyright_stats['Total_copyrights'] else 0
    )

    copyright_table = Table([
        ["Metric", "Value"],
        ["Total Copyrights", copyright_stats['Total_copyrights']],
        ["Success Rate", f"{success_rate:.1f}%"],
        ["Total Cost", f"${copyright_stats['Total_cost_copyright']:,}"]
    ])

    copyright_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    story.append(copyright_table)
    story.append(Spacer(1, 20))
    story.append(Paragraph("A+ Content", section_style))
    story.append(Table([["Total A+", a_plus]]))

    story.append(Spacer(1, 30))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))
    story.append(
        Paragraph(
            f"Generated on {datetime.now().strftime('%B %d, %Y %I:%M %p')}",
            styles['Normal']
        )
    )

    doc.build(story)

    pdf_data = buffer.getvalue()
    buffer.close()

    return pdf_data, filename
