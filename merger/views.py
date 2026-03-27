import json
from django.shortcuts import render, redirect
from django.core.files.base import ContentFile
from .models import MergeJob, UploadedFile
from PyPDF2 import PdfMerger, PdfReader
import csv
import io
import openpyxl
from django.conf import settings
import os
from .utils import docx_to_text, text_to_pdf_buffer

def index(request):
    if request.method == 'POST':
        job_name = request.POST.get('job_name', 'Unnamed Job')
        if not job_name.strip(): job_name = "Merged_Document"
        
        merge_type = request.POST.get('merge_type')
        files = request.FILES.getlist('files')
        file_order = request.POST.get('file_order', '')
        page_ranges_json = request.POST.get('page_ranges', '[]')

        if not files:
            return render(request, 'merger/index.html', {'error': 'No files uploaded.'})

        # Create the job with PENDING status
        job = MergeJob.objects.create(name=job_name, merge_type=merge_type, status='PENDING')

        # Save uploaded files
        uploaded_instances = []
        for f in files:
            uploaded_file = UploadedFile.objects.create(job=job, file=f)
            uploaded_instances.append(uploaded_file)

        # Reorder files based on file_order
        try:
            if file_order:
                order_indices = [int(i) for i in file_order.split(',')]
                valid_indices = [i for i in order_indices if 0 <= i < len(uploaded_instances)]
                ordered_files = [uploaded_instances[i] for i in valid_indices]
                page_ranges = json.loads(page_ranges_json)
            else:
                ordered_files = uploaded_instances
                page_ranges = ['all'] * len(ordered_files)
        except:
            ordered_files = uploaded_instances
            page_ranges = ['all'] * len(ordered_files)

        # Perform merge
        try:
            if merge_type == 'PDF':
                merged_content = merge_pdfs(ordered_files, page_ranges)
                filename = f"{job_name}.pdf"
            elif merge_type == 'CSV':
                merged_content = merge_csvs(ordered_files)
                filename = f"{job_name}.csv"
            elif merge_type == 'EXCEL':
                merged_content = merge_excels(ordered_files)
                filename = f"{job_name}.xlsx"
            else:
                job.status = 'FAILED'
                job.save()
                return render(request, 'merger/index.html', {'error': 'Unsupported merge type.'})

            job.merged_file.save(filename, ContentFile(merged_content))
            job.status = 'COMPLETED'
            job.save()
            return redirect('history')
        except Exception as e:
            # Important: Capture the traceback if debugging
            import traceback
            print(traceback.format_exc())
            job.status = 'FAILED'
            job.save()
            return render(request, 'merger/index.html', {'error': f'Merge failed: {str(e)}'})

    return render(request, 'merger/index.html')

def history(request):
    jobs = MergeJob.objects.all().order_by('-created_at')
    return render(request, 'merger/history.html', {'jobs': jobs})

def parse_page_range(range_str, max_pages):
    if not range_str or range_str.lower() == 'all':
        return range(max_pages)
    pages = []
    for part in range_str.split(','):
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                for p in range(start - 1, end):
                    if 0 <= p < max_pages: pages.append(p)
            except: continue
        else:
            try:
                p = int(part) - 1
                if 0 <= p < max_pages: pages.append(p)
            except: continue
    return pages if pages else range(max_pages)

def merge_pdfs(uploaded_files, page_ranges):
    merger = PdfMerger()
    
    try:
        for i, uf in enumerate(uploaded_files):
            file_path = uf.file.path
            ext = file_path.split('.')[-1].lower()
            range_str = page_ranges[i] if i < len(page_ranges) else 'all'

            if ext == 'pdf':
                with open(file_path, 'rb') as f:
                    reader = PdfReader(f)
                    max_pages = len(reader.pages)
                    pages_to_include = parse_page_range(range_str, max_pages)
                    
                    for p in pages_to_include:
                        merger.append(file_path, pages=(p, p+1))
                        
            elif ext == 'docx':
                try:
                    text = docx_to_text(file_path)
                    if text.strip():
                        docx_pdf_buf = text_to_pdf_buffer(text)
                        merger.append(docx_pdf_buf)
                except Exception as docx_err:
                    print(f"Error merging docx {file_path}: {docx_err}")
                    continue

        output = io.BytesIO()
        merger.write(output)
        return output.getvalue()
    finally:
        merger.close()

def merge_csvs(uploaded_files):
    output = io.StringIO()
    writer = csv.writer(output)
    for i, uf in enumerate(uploaded_files):
        with open(uf.file.path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if i == 0 and header: writer.writerow(header)
            for row in reader: writer.writerow(row)
    return output.getvalue().encode('utf-8')

def merge_excels(uploaded_files):
    new_wb = openpyxl.Workbook()
    default_sheet = new_wb.active
    new_wb.remove(default_sheet)
    for uf in uploaded_files:
        wb = openpyxl.load_workbook(uf.file.path, data_only=True)
        for sheet_name in wb.sheetnames:
            source_sheet = wb[sheet_name]
            target_sheet_name = f"{os.path.basename(uf.file.name)}_{sheet_name}"[:31]
            target_sheet = new_wb.create_sheet(title=target_sheet_name)
            for row in source_sheet.iter_rows(values_only=True):
                target_sheet.append(row)
    output = io.BytesIO()
    new_wb.save(output)
    return output.getvalue()
