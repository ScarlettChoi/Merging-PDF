import os
import PyPDF2

def split_pdfs(folder_path, output_folder):
    for root, _, files in os.walk(folder_path):
        pdf_files = [os.path.join(root, file) for file in files if file.lower().endswith('.pdf')]
        for pdf in sorted(pdf_files):
            with open(pdf, 'rb') as infile:
                reader = PyPDF2.PdfReader(infile)
                for i in range(len(reader.pages)):
                    writer = PyPDF2.PdfWriter()
                    writer.add_page(reader.pages[i])
                    output_filename = os.path.join(output_folder, f'{os.path.splitext(os.path.basename(pdf))[0]}_page_{i+1}.pdf')
                    with open(output_filename, 'wb') as outfile:
                        writer.write(outfile)
                        print(f'Saved: {output_filename}')

if __name__ == "__main__":
    folder_path = 'PDF for Split'  # PDF 파일이 있는 폴더의 상대 경로
    output_folder = 'Split Files'  # 분할된 PDF 파일을 저장할 출력 폴더의 상대 경로

    # 출력 폴더가 존재하지 않으면 생성
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    split_pdfs(folder_path, output_folder)
