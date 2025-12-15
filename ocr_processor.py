import streamlit as st
import requests
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter
import io
import os
import base64
import numpy as np

class OCRProcessor:
    def __init__(self):
        """Initialize Gemini AI for OCR processing"""
        # Try Streamlit Cloud secrets first, then env var
        api_key = None
        try:
            api_key = st.secrets.get("GEMINI_API_KEY", None)  # type: ignore[attr-defined]
        except Exception:
            api_key = None
        if not api_key:
            api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable is required")
        
        self.api_key = api_key
        # Gemini REST API endpoint (multimodal)
        self.model_name = 'Gemini 2.5 Flash-Lite'
        self.endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model_name}:generateContent"
        
        # OCR prompt for Japanese handwritten text
        self.ocr_prompt = """
あなたは日本語手書き文字認識の専門家です。画像を非常に注意深く観察し、書かれた文字を一字一句正確に読み取ってください。

**超精密文字認識手順：**

1. **画像全体の観察**
   - 全体レイアウトを把握
   - 質問と回答の位置関係を確認
   - 文字の配置パターンを理解

2. **文字単位での詳細分析**
   - 各文字の形状を慎重に観察
   - 画線の太さ、角度、曲がり具合を分析
   - 文字の一部が欠けていても、見える部分から判断

3. **日本語文字の特別な注意点**
   - ひらがな：「る/ろ」「は/ば/ぱ」「き/さ」「な/た」「わ/れ」「め/ぬ」
   - カタカナ：「ソ/ン」「シ/ツ」「ク/ワ」「ロ/コ」「エ/ユ」
   - 漢字：似た形の字に特に注意
   - 濁点・半濁点の有無を慎重に確認

4. **文字起こしの原則**
   - 書かれている通りに正確に転写
   - 推測や修正は一切しない
   - 誤字があってもそのまま記録
   - 判読不能な文字は「？」で表記

5. **出力形式**
   - 質問：Q1: [質問文そのまま]
   - 回答：質問の直下に書かれた通りに記載
   - 空欄：（無回答）

**絶対に守ること：**
- 見えた文字をそのまま書く
- 意味を考えて修正しない
- 推測で文字を補完しない

この画像の手書き文字を正確に読み取ってください：
"""

    def process_pdf(self, pdf_path):
        """Process PDF file and extract handwritten Japanese text using Gemini AI"""
        try:
            # Open PDF
            pdf_document = fitz.open(pdf_path)
            all_images = []
            
            # First, collect all page images
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                
                # Convert page to image with maximum resolution
                mat = fitz.Matrix(6.0, 6.0)  # Maximum resolution for best OCR
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                
                # Convert to PIL Image and apply multiple enhancement techniques
                image = Image.open(io.BytesIO(img_data))
                
                # Apply advanced image enhancement for maximum OCR accuracy
                image = self._enhance_image_for_ocr(image)
                image = self._apply_additional_preprocessing(image)
                
                all_images.append((image, page_num + 1))
            
            pdf_document.close()
            
            # Try individual processing first, fallback to batch if needed
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"PDF処理エラー: {str(e)}")
    
    def process_image(self, image_path):
        """Process image file (PNG, JPG, etc.) and extract handwritten Japanese text using Gemini AI"""
        try:
            # Open image file
            image = Image.open(image_path)
            
            # Convert to RGB if necessary (handle RGBA, P mode, etc.)
            if image.mode in ('RGBA', 'P', 'LA'):
                # Create white background for transparent images
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None)
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Apply image enhancement for OCR
            image = self._enhance_image_for_ocr(image)
            image = self._apply_additional_preprocessing(image)
            
            # Process as single image
            all_images = [(image, 1)]
            
            # Try individual processing
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")
    
    def process_images(self, image_paths):
        """Process multiple image files and extract handwritten Japanese text using Gemini AI"""
        try:
            all_images = []
            
            for idx, image_path in enumerate(image_paths):
                # Open image file
                image = Image.open(image_path)
                
                # Convert to RGB if necessary
                if image.mode in ('RGBA', 'P', 'LA'):
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    if image.mode == 'P':
                        image = image.convert('RGBA')
                    background.paste(image, mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None)
                    image = background
                elif image.mode != 'RGB':
                    image = image.convert('RGB')
                
                # Apply image enhancement for OCR
                image = self._enhance_image_for_ocr(image)
                image = self._apply_additional_preprocessing(image)
                
                all_images.append((image, idx + 1))
            
            # Try individual processing first, fallback to batch if needed
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")
    
    def process_image_bytes(self, image_bytes, filename="image"):
        """Process image from bytes data"""
        try:
            # Open image from bytes
            image = Image.open(io.BytesIO(image_bytes))
            
            # Convert to RGB if necessary
            if image.mode in ('RGBA', 'P', 'LA'):
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None)
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Apply image enhancement for OCR
            image = self._enhance_image_for_ocr(image)
            image = self._apply_additional_preprocessing(image)
            
            # Process as single image
            all_images = [(image, 1)]
            
            # Try individual processing
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")
    
    def _enhance_image_for_ocr(self, image):
        """Enhance image quality for better OCR accuracy"""
        try:
            # Convert to RGB if necessary
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 1. Increase contrast for better text visibility
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(1.3)
            
            # 2. Increase sharpness to make text edges clearer
            enhancer = ImageEnhance.Sharpness(image)
            image = enhancer.enhance(1.5)
            
            # 3. Adjust brightness if too dark
            enhancer = ImageEnhance.Brightness(image)
            image = enhancer.enhance(1.1)
            
            # 4. Apply slight noise reduction
            image = image.filter(ImageFilter.MedianFilter(size=1))
            
            # 5. Apply unsharp mask for text clarity
            image = image.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=3))
            
            return image
            
        except Exception as e:
            # Return original image if enhancement fails
            return image
    
    def _apply_additional_preprocessing(self, image):
        """Apply advanced preprocessing techniques for maximum OCR accuracy"""
        try:
            # Convert to grayscale for better text recognition
            if image.mode != 'L':
                image = image.convert('L')
            
            # Convert to numpy array for advanced processing
            img_array = np.array(image)
            
            # Apply adaptive threshold to improve text contrast
            from PIL import ImageOps
            image = Image.fromarray(img_array)
            
            # Increase contrast more aggressively
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(2.0)
            
            # Apply edge enhancement
            image = image.filter(ImageFilter.EDGE_ENHANCE_MORE)
            
            # Convert back to RGB for Gemini compatibility
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            return image
            
        except Exception as e:
            # Return original image if preprocessing fails
            return image
    
    def _extract_text_from_image(self, image, page_num):
        """Extract text from image using Gemini AI"""
        try:
            # Convert PIL Image to bytes
            img_byte_arr = io.BytesIO()
            image.save(img_byte_arr, format='PNG')
            img_byte_arr = img_byte_arr.getvalue()
            
            # Create Gemini compatible image
            gemini_image = {
                "mime_type": "image/png",
                "data": img_byte_arr
            }
            
            # Generate content with Gemini
            headers = {
                "Content-Type": "application/json"
            }
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": self.ocr_prompt},
                            {
                                "inline_data": {
                                    "mime_type": "image/png",
                                    "data": base64.b64encode(img_byte_arr).decode("utf-8")
                                }
                            }
                        ]
                    }
                ]
            }
            res = requests.post(
                f"{self.endpoint}?key={self.api_key}",
                headers=headers,
                json=payload,
                timeout=120
            )
            res.raise_for_status()
            data = res.json()
            text = (
                data.get("candidates", [{}])[0]
                .get("content", {})
                .get("parts", [{}])[0]
                .get("text", "")
            )
            if text:
                return text.strip()
            else:
                return f"ページ {page_num}: テキストが検出されませんでした"
                
        except Exception as e:
            return f"ページ {page_num}: OCRエラー - {str(e)}"
    
    def _extract_text_from_all_images(self, images_with_pages):
        """Extract text from multiple images with question consolidation"""
        try:
            # Enhanced multi-page processing prompt
            multi_prompt = f"""
{self.ocr_prompt}

**マルチページ処理特別指示：**
現在、{len(images_with_pages)}ページのアンケートを同時処理しています。

**統合ルール（厳格遵守）：**
1. **質問の統合**：同一質問が複数ページに出現する場合、質問文は一度のみ記載
2. **回答の統合**：各ページの回答を「・」で区切り、回答ごとに箇条書きにして改行する
3. **文字起こし原則**：書かれている文字をそのまま正確に転写（意訳禁止）
4. **重複除去しない**：同じ回答でも個別に記載
5. **論理的順序**：質問番号順（Q1、Q2...）で整理

**期待される統合例：**
ページ1: Q1: 満足度は？ → とても満足
ページ1: Q2: 理解度は？ → よく理解できた
ページ2: Q1: 満足度は？ → 満足
ページ2: Q2: 理解度は？ → 理解できなかった
ページ3: Q1: 満足度は？ → 普通
ページ3: Q2: 理解度は？ → 理解できた

↓ 統合後 ↓
Q1: 満足度は？
・とても満足
・満足
・普通

Q2: 理解度は？
・よく理解できた
・理解できなかった
・理解できた

**品質要求：**
- 各文字を一字一句正確に認識
- 筆跡の個人差を考慮した推論
- 書かれている文字をそのまま転写（修正・意訳禁止）
"""
            
            # Create content parts list
            content_parts = [multi_prompt]
            
            # Add all images
            for image, page_num in images_with_pages:
                # Convert PIL Image to bytes
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                # Create Gemini compatible image
                gemini_image = {
                    "mime_type": "image/png",
                    "data": img_byte_arr
                }
                
                content_parts.append(f"\n--- ページ {page_num} ---")
                content_parts.append(gemini_image)
            
            # Build payload for multi-page
            parts = []
            for part in content_parts:
                if isinstance(part, str):
                    parts.append({"text": part})
                else:
                    # image dict with bytes
                    parts.append({
                        "inline_data": {
                            "mime_type": part.get("mime_type", "image/png"),
                            "data": base64.b64encode(part["data"]).decode("utf-8")
                        }
                    })
            payload = {"contents": [{"parts": parts}]}
            res = requests.post(
                f"{self.endpoint}?key={self.api_key}",
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=300
            )
            res.raise_for_status()
            data = res.json()
            text = (
                data.get("candidates", [{}])[0]
                .get("content", {})
                .get("parts", [{}])[0]
                .get("text", "")
            )
            if text:
                return text.strip()
            else:
                return "テキストが検出されませんでした"
                
        except Exception as e:
            # Fallback to individual page processing if multi-image fails
            all_text = []
            for image, page_num in images_with_pages:
                page_text = self._extract_text_from_image(image, page_num)
                if page_text and page_text.strip():
                    all_text.append(f"--- ページ {page_num} ---\n{page_text}")
            
            return "\n\n".join(all_text) if all_text else "文字起こし結果がありません。"
    
    def _process_pages_individually(self, images_with_pages):
        """Process each page individually with retry logic for maximum accuracy"""
        all_page_results = []
        
        for image, page_num in images_with_pages:
            # Try multiple extraction attempts for each page
            page_text = self._extract_with_retry(image, page_num)
            if page_text and page_text.strip():
                all_page_results.append({
                    'page_num': page_num,
                    'text': page_text.strip()
                })
        
        # Consolidate questions across pages
        return self._consolidate_questions_from_pages(all_page_results)
    
    def _extract_with_retry(self, image, page_num, max_retries=1):
        """Extract text with multiple attempts for better accuracy"""
        best_result = ""
        
        for attempt in range(max_retries):
            try:
                # Convert PIL Image to bytes with high quality
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG', quality=100, optimize=False)
                img_byte_arr = img_byte_arr.getvalue()
                
                # Enhanced prompt for individual page processing
                individual_prompt = f"""
{self.ocr_prompt}

**ページ {page_num} の詳細分析：**
- この1ページのみに集中してください
- 書かれた文字を一字一句正確に読み取ってください
- 質問番号（Q1、Q2など）があれば必ず含めてください
- 各質問の回答を正確に記録してください

画像を慎重に分析し、手書き文字を正確に文字起こししてください：
"""
                
                # Create image content for Gemini
                content = [
                    individual_prompt,
                    {
                        "mime_type": "image/png", 
                        "data": img_byte_arr
                    }
                ]
                
                # REST call
                payload = {
                    "contents": [
                        {
                            "parts": [
                                {"text": individual_prompt},
                                {
                                    "inline_data": {
                                        "mime_type": "image/png",
                                        "data": base64.b64encode(img_byte_arr).decode("utf-8")
                                    }
                                }
                            ]
                        }
                    ]
                }
                res = requests.post(
                    f"{self.endpoint}?key={self.api_key}",
                    headers={"Content-Type": "application/json"},
                    json=payload,
                    timeout=120
                )
                res.raise_for_status()
                data = res.json()
                current_result = (
                    data.get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
                ).strip()
                if current_result:
                    
                    # Keep the longest/most detailed result
                    if len(current_result) > len(best_result):
                        best_result = current_result
                
            except Exception as e:
                error_message = str(e)
                print(f"Attempt {attempt + 1} failed for page {page_num}: {e}")
                
                # If quota exceeded, wait and don't retry immediately
                if "429" in error_message or "quota" in error_message.lower():
                    print(f"Quota exceeded for page {page_num}, skipping retries")
                    break
                    
                continue
        
        return best_result if best_result else f"ページ {page_num}: 文字起こしに失敗しました"
    
    def _consolidate_questions_from_pages(self, page_results):
        """Consolidate questions and answers from multiple pages"""
        questions_dict = {}
        
        # Parse each page's results
        for page_result in page_results:
            page_num = page_result['page_num']
            text = page_result['text']
            
            # Extract questions and answers from this page
            questions = self._parse_questions_from_text(text)
            
            for q_num, q_text, answer in questions:
                if q_num not in questions_dict:
                    questions_dict[q_num] = {
                        'question': q_text,
                        'answers': []
                    }
                
                # Add answer if it's not empty
                if answer and answer.strip() and answer.strip() not in ['（無回答）', '']:
                    questions_dict[q_num]['answers'].append(answer.strip())
        
        # Build final consolidated text
        result_lines = []
        for q_num in sorted(questions_dict.keys()):
            q_data = questions_dict[q_num]
            result_lines.append(f"Q{q_num}: {q_data['question']}")
            
            # Add answers as bullet points
            for answer in q_data['answers']:
                result_lines.append(f"・{answer}")
            
            result_lines.append("")  # Empty line between questions
        
        return "\n".join(result_lines)
    
    def _parse_questions_from_text(self, text):
        """Parse questions and answers from text"""
        questions = []
        lines = text.split('\n')
        current_question = None
        current_q_num = None
        current_answer = ""
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check if line is a question
            import re
            q_match = re.match(r'^Q(\d+)[:：]\s*(.+)', line)
            if q_match:
                # Save previous question if exists
                if current_question and current_q_num:
                    questions.append((current_q_num, current_question, current_answer.strip()))
                
                # Start new question
                current_q_num = int(q_match.group(1))
                current_question = q_match.group(2).strip()
                current_answer = ""
            else:
                # This is likely an answer
                if current_question:
                    if current_answer:
                        current_answer += " " + line
                    else:
                        current_answer = line
        
        # Add the last question
        if current_question and current_q_num:
            questions.append((current_q_num, current_question, current_answer.strip()))
        
        return questions
    
    def _post_process_text(self, text):
        """Post-process extracted text for better readability"""
        if not text:
            return "文字起こし結果がありません。"
        
        # Clean up common OCR artifacts
        lines = text.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            if line:  # Skip empty lines
                # Remove "回答：" prefix if present
                if line.startswith('回答：'):
                    line = line[3:].strip()
                elif line.startswith('回答:'):
                    line = line[3:].strip()
                
                # Normalize spaces - remove irregular half-width/full-width spaces
                import re
                # Replace multiple spaces with single space
                line = re.sub(r'\s+', ' ', line)
                # Remove spaces around Japanese punctuation
                line = re.sub(r'\s*([。、！？：；])\s*', r'\1', line)
                # Remove unnecessary spaces in Japanese text
                line = re.sub(r'([あ-ん])\s+([あ-ん])', r'\1\2', line)
                line = re.sub(r'([ア-ン])\s+([ア-ン])', r'\1\2', line)
                line = re.sub(r'([一-龯])\s+([一-龯])', r'\1\2', line)
                
                cleaned_lines.append(line)
        
        # Join lines with proper spacing
        result = '\n'.join(cleaned_lines)
        
        # Add final processing note
        if result:
            result += "\n\n--- 注意 ---\n手書き文字の自動認識結果です。不正確な部分がある可能性があります。"
        
        return result if result else "文字起こし結果がありません。"
