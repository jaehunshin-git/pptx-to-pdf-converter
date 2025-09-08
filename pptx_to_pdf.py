import os
import comtypes.client

def convert_all_ppt_to_pdf(input_dir: str, output_dir: str):
    """
    입력 디렉토리 내의 모든 PPT/PPTX 파일을 PDF로 변환하여 출력 디렉토리에 저장

    :param input_dir: PPT/PPTX 파일이 들어있는 폴더 경로
    :param output_dir: PDF 파일을 저장할 폴더 경로
    """
    # 절대경로로 변환
    input_dir = os.path.abspath(input_dir)
    output_dir = os.path.abspath(output_dir)

    # 출력 디렉토리가 없으면 생성
    os.makedirs(output_dir, exist_ok=True)

    # PowerPoint COM 객체 생성
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1  # 백그라운드 실행

    try:
        for filename in os.listdir(input_dir):
            if filename.lower().endswith((".pptx", ".ppt")):
                ppt_path = os.path.join(input_dir, filename)
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = os.path.join(output_dir, pdf_name)

                print(f"변환 중: {ppt_path} -> {pdf_path}")
                
                # 프레젠테이션 열고 PDF로 저장
                presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
                presentation.SaveAs(pdf_path, FileFormat=32)  # 32 = PDF
                presentation.Close()
    finally:
        powerpoint.Quit()
        print("모든 변환 완료.")

# 예시 실행
if __name__ == "__main__":
    input_folder = r"C:\Users\JHSHIN\ProgrammingCodes\solo-project\pdf-related\pptxtopdf\input_dir"
    output_folder = r"C:\Users\JHSHIN\ProgrammingCodes\solo-project\pdf-related\pptxtopdf\output_dir"
    convert_all_ppt_to_pdf(input_folder, output_folder)
