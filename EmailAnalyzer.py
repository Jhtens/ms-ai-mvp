import streamlit as st
import openai
from openai import AzureOpenAI
import json
from datetime import datetime, timedelta
import re
from typing import Dict, List, Any
import email
import email.utils
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import io
import base64
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import zipfile
import tempfile
import os
import platform

# 페이지 설정
st.set_page_config(
    page_title="E-mail Analyzer",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)


class EmailParser:
    """이메일 파일 파싱 클래스"""

    @staticmethod
    def parse_eml_file(file_content: bytes) -> Dict[str, str]:
        """EML 파일 파싱"""
        try:
            msg = email.message_from_bytes(file_content)

            # 헤더 정보 추출
            subject = msg.get("Subject", "")
            sender = msg.get("From", "")
            recipients = msg.get("To", "")
            date = msg.get("Date", "")

            # 본문 추출
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body += part.get_payload(decode=True).decode(
                            "utf-8", errors="ignore"
                        )
                    elif part.get_content_type() == "text/html" and not body:
                        # HTML이지만 plain text가 없는 경우
                        html_content = part.get_payload(decode=True).decode(
                            "utf-8", errors="ignore"
                        )
                        # 간단한 HTML 태그 제거
                        body = re.sub(r"<[^>]+>", "", html_content)
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            return {
                "subject": subject,
                "sender": sender,
                "recipients": recipients,
                "date": date,
                "body": body,
                "full_content": f"제목: {subject}\n발신자: {sender}\n수신자: {recipients}\n날짜: {date}\n\n{body}",
            }

        except Exception as e:
            st.error(f"EMail file parsing error: {str(e)}")
            return None

    @staticmethod
    def parse_msg_file(file_content: bytes) -> Dict[str, str]:
        """MSG 파일 파싱 (기본적인 구현)"""
        try:
            # MSG 파일은 복잡한 바이너리 형식이므로 간단한 텍스트 추출만 구현
            content = file_content.decode("utf-8", errors="ignore")

            # 기본적인 정보만 추출
            return {
                "subject": "파싱된 MSG 파일",
                "sender": "알 수 없음",
                "recipients": "알 수 없음",
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "body": content,
                "full_content": content,
            }

        except Exception as e:
            st.error(f"MSG 파일 파싱 오류: {str(e)}")
            return None


class CalendarIntegration:
    """일정 연동 클래스"""

    @staticmethod
    def create_ics_event(task: Dict[str, Any], email_subject: str = "") -> str:
        """ICS 형식의 일정 생성"""
        try:
            event_id = f"email-task-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            now = datetime.now().strftime("%Y%m%dT%H%M%SZ")

            # 마감일 처리
            if task.get("deadline") and task["deadline"] != "null":
                # 간단한 날짜 파싱
                deadline_str = task["deadline"]
                try:
                    # 여러 날짜 형식 시도
                    for fmt in ["%Y-%m-%d", "%m/%d", "%m월 %d일", "%d일"]:
                        try:
                            if fmt in ["%m/%d", "%m월 %d일", "%d일"]:
                                # 현재 연도 추가
                                if "월" in deadline_str:
                                    deadline_str = (
                                        f"{datetime.now().year}년 {deadline_str}"
                                    )
                                    due_date = datetime.strptime(
                                        deadline_str, f"%Y년 %m월 %d일"
                                    )
                                elif "일" in deadline_str:
                                    due_date = datetime.now() + timedelta(
                                        days=int(deadline_str.replace("일", ""))
                                    )
                                else:
                                    due_date = datetime.strptime(
                                        f"{datetime.now().year}/{deadline_str}",
                                        f"%Y/{fmt}",
                                    )
                            else:
                                due_date = datetime.strptime(deadline_str, fmt)
                            break
                        except ValueError:
                            continue
                    else:
                        # 파싱 실패 시 1주일 후로 설정
                        due_date = datetime.now() + timedelta(days=7)
                except:
                    due_date = datetime.now() + timedelta(days=7)
            else:
                due_date = datetime.now() + timedelta(days=7)

            due_date_str = due_date.strftime("%Y%m%dT%H%M%SZ")

            ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Email Analyzer//Task//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VEVENT
UID:{event_id}
DTSTART:{due_date_str}
DTEND:{due_date_str}
DTSTAMP:{now}
SUMMARY:{task.get('task', '할일')}
DESCRIPTION:우선순위: {task.get('priority', 'medium')}\n담당자: {task.get('assignee', '미정')}\n관련 이메일: {email_subject}
PRIORITY:{5 if task.get('priority') == 'low' else 1 if task.get('priority') == 'high' else 3}
STATUS:NEEDS-ACTION
END:VEVENT
END:VCALENDAR"""

            return ics_content

        except Exception as e:
            st.error(f"일정 생성 오류: {str(e)}")
            return None

    @staticmethod
    def create_calendar_summary(tasks: List[Dict[str, Any]]) -> pd.DataFrame:
        """일정 요약 테이블 생성"""
        calendar_data = []

        for i, task in enumerate(tasks, 1):
            # 마감일 처리
            if task.get("deadline") and task["deadline"] != "null":
                deadline = task["deadline"]
            else:
                deadline = "미정"

            calendar_data.append(
                {
                    "번호": i,
                    "할일": task.get("task", ""),
                    "우선순위": task.get("priority", "medium"),
                    "마감일": deadline,
                    "담당자": task.get("assignee", "미정"),
                    "상태": "미완료",
                }
            )

        return pd.DataFrame(calendar_data)


class ExportManager:
    """결과 내보내기 관리 클래스"""

    @staticmethod
    def download_and_register_font():
        """나눔고딕 폰트를 다운로드하고 등록합니다."""
        font_dir = "fonts"
        font_path = os.path.join(font_dir, "NanumGothic.ttf")
        font_url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"

        if not os.path.exists(font_path):
            try:
                os.makedirs(font_dir, exist_ok=True)
                with st.spinner(f"한글 폰트를 다운로드합니다... ({font_path})"):
                    import urllib.request

                    urllib.request.urlretrieve(font_url, font_path)
            except Exception as e:
                st.error(f"폰트 다운로드 실패: {e}")
                return "Helvetica"

        try:
            pdfmetrics.registerFont(TTFont("Korean", font_path))
            return "Korean"
        except Exception as e:
            st.error(f"폰트 등록 실패: {e}")
            return "Helvetica"

    @staticmethod
    def setup_korean_fonts():
        """한글 폰트 설정 (나눔고딕 다운로드 방식)"""
        try:
            # 이미 등록되었는지 확인
            pdfmetrics.getFont("Korean")
            return "Korean"
        except KeyError:
            # 등록되지 않았다면 다운로드 및 등록
            return ExportManager.download_and_register_font()

    @staticmethod
    def create_korean_styles(font_name="Korean"):
        """한글 지원 스타일 생성"""
        styles = getSampleStyleSheet()

        # 한글 제목 스타일
        korean_title = ParagraphStyle(
            "KoreanTitle",
            parent=styles["Title"],
            fontName=font_name,
            fontSize=20,
            spaceAfter=30,
            textColor=colors.darkblue,
            alignment=TA_CENTER,
        )

        # 한글 제목2 스타일
        korean_heading2 = ParagraphStyle(
            "KoreanHeading2",
            parent=styles["Heading2"],
            fontName=font_name,
            fontSize=16,
            spaceAfter=12,
            spaceBefore=12,
            textColor=colors.darkblue,
        )

        # 한글 일반 텍스트 스타일
        korean_normal = ParagraphStyle(
            "KoreanNormal",
            parent=styles["Normal"],
            fontName=font_name,
            fontSize=10,
            spaceAfter=6,
            leading=14,
        )

        # 한글 작은 텍스트 스타일
        korean_small = ParagraphStyle(
            "KoreanSmall",
            parent=styles["Normal"],
            fontName=font_name,
            fontSize=8,
            spaceAfter=4,
            leading=10,
        )

        return {
            "title": korean_title,
            "heading2": korean_heading2,
            "normal": korean_normal,
            "small": korean_small,
        }

    @staticmethod
    def create_pdf_report(
        analysis_result: Dict[str, Any], email_content: str = ""
    ) -> bytes:
        """PDF 보고서 생성 (한글 지원)"""
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1 * inch)
        story = []

        try:
            # 한글 폰트 설정
            font_name = ExportManager.setup_korean_fonts()
            korean_styles = ExportManager.create_korean_styles(font_name)

            # 제목
            story.append(Paragraph("이메일 분석 보고서", korean_styles["title"]))
            story.append(Spacer(1, 20))

            # 분석 일시
            story.append(
                Paragraph(
                    f"분석 일시: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M')}",
                    korean_styles["normal"],
                )
            )
            story.append(Spacer(1, 20))

            # 요약
            story.append(Paragraph("📝 요약", korean_styles["heading2"]))
            summary_text = str(analysis_result.get("summary", "요약 정보 없음"))
            # HTML 태그 제거 및 특수문자 처리
            summary_text = (
                summary_text.replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("&", "&amp;")
            )
            story.append(Paragraph(summary_text, korean_styles["normal"]))
            story.append(Spacer(1, 15))

            # 긴급도 및 감정
            story.append(Paragraph("📊 분석 지표", korean_styles["heading2"]))
            indicators_data = [
                ["지표", "결과"],
                ["긴급도", str(analysis_result.get("urgency_level", "medium")).upper()],
                ["감정 톤", str(analysis_result.get("sentiment", "neutral")).upper()],
            ]

            indicators_table = Table(indicators_data, colWidths=[2 * inch, 2 * inch])
            indicators_table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("FONTNAME", (0, 0), (-1, -1), font_name),
                        ("FONTSIZE", (0, 0), (-1, 0), 12),
                        ("FONTSIZE", (0, 1), (-1, -1), 10),
                        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                        ("GRID", (0, 0), (-1, -1), 1, colors.black),
                    ]
                )
            )
            story.append(indicators_table)
            story.append(Spacer(1, 20))

            # 할일 목록
            if analysis_result.get("tasks") and len(analysis_result["tasks"]) > 0:
                story.append(Paragraph("✅ 할일 목록", korean_styles["heading2"]))
                tasks_data = [["번호", "할일", "우선순위", "마감일", "담당자"]]

                for i, task in enumerate(analysis_result["tasks"], 1):
                    task_text = str(task.get("task", ""))
                    if len(task_text) > 40:
                        task_text = task_text[:40] + "..."

                    # 특수문자 처리
                    task_text = (
                        task_text.replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace("&", "&amp;")
                    )

                    deadline = task.get("deadline", "미정")
                    if deadline == "null" or deadline is None:
                        deadline = "미정"

                    assignee = task.get("assignee", "미정")
                    if assignee == "null" or assignee is None:
                        assignee = "미정"

                    tasks_data.append(
                        [
                            str(i),
                            task_text,
                            str(task.get("priority", "medium")),
                            str(deadline),
                            str(assignee),
                        ]
                    )

                tasks_table = Table(
                    tasks_data,
                    colWidths=[
                        0.4 * inch,
                        2.6 * inch,
                        0.8 * inch,
                        1 * inch,
                        0.8 * inch,
                    ],
                )
                tasks_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("FONTNAME", (0, 0), (-1, -1), font_name),
                            ("FONTSIZE", (0, 0), (-1, 0), 9),
                            ("FONTSIZE", (0, 1), (-1, -1), 8),
                            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                            ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                            ("GRID", (0, 0), (-1, -1), 1, colors.black),
                            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ]
                    )
                )
                story.append(tasks_table)
                story.append(Spacer(1, 20))

            # 주요 포인트
            if (
                analysis_result.get("key_points")
                and len(analysis_result["key_points"]) > 0
            ):
                story.append(Paragraph("🎯 주요 포인트", korean_styles["heading2"]))
                for i, point in enumerate(analysis_result["key_points"], 1):
                    point_text = (
                        str(point)
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace("&", "&amp;")
                    )
                    story.append(
                        Paragraph(f"{i}. {point_text}", korean_styles["normal"])
                    )
                story.append(Spacer(1, 15))

            # 즉시 처리 항목
            if (
                analysis_result.get("action_items")
                and len(analysis_result["action_items"]) > 0
            ):
                story.append(Paragraph("⚡ 즉시 처리 항목", korean_styles["heading2"]))
                for i, item in enumerate(analysis_result["action_items"], 1):
                    item_text = (
                        str(item)
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace("&", "&amp;")
                    )
                    story.append(
                        Paragraph(f"🚨 {i}. {item_text}", korean_styles["normal"])
                    )
                story.append(Spacer(1, 15))

            # 후속 조치
            if analysis_result.get("follow_up"):
                story.append(Paragraph("🔄 후속 조치", korean_styles["heading2"]))
                followup_text = (
                    str(analysis_result["follow_up"])
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("&", "&amp;")
                )
                story.append(Paragraph(followup_text, korean_styles["normal"]))

            # PDF 생성
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

        except Exception as e:
            # 오류 발생 시 기본 영문 PDF 생성
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            story = []

            styles = getSampleStyleSheet()
            story.append(
                Paragraph(
                    "Email Analysis Report (Error in Korean PDF)", styles["Title"]
                )
            )
            story.append(Spacer(1, 20))
            story.append(Paragraph(f"Error occurred: {str(e)}", styles["Normal"]))
            story.append(Spacer(1, 10))
            story.append(
                Paragraph(
                    f"Analysis Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    styles["Normal"],
                )
            )

            # 기본 정보라도 포함
            if analysis_result.get("summary"):
                story.append(Spacer(1, 20))
                story.append(Paragraph("Summary:", styles["Heading2"]))
                story.append(
                    Paragraph(str(analysis_result["summary"]), styles["Normal"])
                )

            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()

    @staticmethod
    def create_excel_report(analysis_result: Dict[str, Any]) -> bytes:
        """Excel 보고서 생성"""
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # 요약 시트
            summary_data = {
                "항목": ["요약", "긴급도", "감정 톤", "분석 일시"],
                "내용": [
                    analysis_result.get("summary", ""),
                    analysis_result.get("urgency_level", "medium"),
                    analysis_result.get("sentiment", "neutral"),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ],
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="요약", index=False)

            # 할일 목록 시트
            if analysis_result.get("tasks"):
                tasks_data = []
                for i, task in enumerate(analysis_result["tasks"], 1):
                    tasks_data.append(
                        {
                            "번호": i,
                            "할일": task.get("task", ""),
                            "우선순위": task.get("priority", "medium"),
                            "마감일": (
                                task.get("deadline", "미정")
                                if task.get("deadline") != "null"
                                else "미정"
                            ),
                            "담당자": (
                                task.get("assignee", "미정")
                                if task.get("assignee") != "null"
                                else "미정"
                            ),
                            "상태": "미완료",
                        }
                    )

                tasks_df = pd.DataFrame(tasks_data)
                tasks_df.to_excel(writer, sheet_name="할일목록", index=False)

            # 주요 포인트 시트
            if analysis_result.get("key_points"):
                points_data = {
                    "번호": range(1, len(analysis_result["key_points"]) + 1),
                    "주요 포인트": analysis_result["key_points"],
                }
                points_df = pd.DataFrame(points_data)
                points_df.to_excel(writer, sheet_name="주요포인트", index=False)

        buffer.seek(0)
        return buffer.getvalue()

    @staticmethod
    def create_calendar_zip(
        tasks: List[Dict[str, Any]], email_subject: str = ""
    ) -> bytes:
        """일정 파일들을 ZIP으로 압축"""
        buffer = io.BytesIO()

        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, task in enumerate(tasks, 1):
                ics_content = CalendarIntegration.create_ics_event(task, email_subject)
                if ics_content:
                    filename = f"task_{i}_{task.get('task', 'task')[:20]}.ics"
                    # 파일명에서 특수문자 제거
                    filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
                    zip_file.writestr(filename, ics_content)

        buffer.seek(0)
        return buffer.getvalue()


class EmailAnalyzer:
    def __init__(self):
        self.client = None

    def initialize_azure_openai(
        self, endpoint: str, api_key: str, api_version: str = "2023-12-01-preview"
    ):
        """Azure OpenAI 클라이언트 초기화"""
        try:
            self.client = AzureOpenAI(
                azure_endpoint=endpoint, api_key=api_key, api_version=api_version
            )
            return True
        except Exception as e:
            st.error(f"Azure OpenAI 연결 실패: {str(e)}")
            return False

    def analyze_email(
        self, email_content: str, model_name: str = "gpt-4"
    ) -> Dict[str, Any]:
        """이메일 내용을 분석하여 업무 요약과 할일 추출"""
        if not self.client:
            raise Exception("Azure OpenAI 클라이언트가 초기화되지 않았습니다.")

        system_prompt = """
        당신은 이메일을 분석해서 업무 요약과 할일을 추출하는 전문가입니다.
        주어진 이메일 내용을 분석해서 다음 형식의 JSON으로 응답하세요:

        {
            "summary": "이메일의 주요 내용을 2-3문장으로 요약",
            "key_points": ["주요 포인트1", "주요 포인트2", "주요 포인트3"],
            "tasks": [
                {
                    "task": "할일 내용",
                    "priority": "high|medium|low",
                    "deadline": "마감일 (YYYY-MM-DD 형식 또는 상대적 표현)",
                    "assignee": "담당자 (없으면 null)"
                }
            ],
            "action_items": ["즉시 처리해야 할 항목1", "즉시 처리해야 할 항목2"],
            "follow_up": "후속 조치가 필요한 사항",
            "sentiment": "positive|neutral|negative",
            "urgency_level": "high|medium|low"
        }

        한국어로 분석하고, 실제 업무에 도움이 되는 구체적인 할일을 추출하세요.
        마감일은 가능한 구체적으로 추출하되, 명시되지 않았으면 null로 설정하세요.
        """

        try:
            response = self.client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": f"다음 이메일을 분석해주세요:\n\n{email_content}",
                    },
                ],
                temperature=0.3,
                max_tokens=2000,
            )

            # JSON 응답 파싱
            response_text = response.choices[0].message.content.strip()

            # JSON 블록 추출
            json_match = re.search(r"```json\n(.*?)\n```", response_text, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                json_str = response_text

            try:
                analysis_result = json.loads(json_str)
            except json.JSONDecodeError:
                # JSON 파싱 실패 시 재시도
                analysis_result = self._parse_fallback_response(response_text)

            return analysis_result

        except Exception as e:
            st.error(f"분석 중 오류 발생: {str(e)}")
            return None

    def _parse_fallback_response(self, response_text: str) -> Dict[str, Any]:
        """JSON 파싱 실패 시 대체 파싱 방법"""
        return {
            "summary": "분석 결과를 파싱하는 중 오류가 발생했습니다.",
            "key_points": ["원시 응답을 확인하세요."],
            "tasks": [],
            "action_items": [],
            "follow_up": "다시 시도해주세요.",
            "sentiment": "neutral",
            "urgency_level": "medium",
            "raw_response": response_text,
        }


def main():
    st.title("📧 E-mail Analyzer")
    st.markdown("---")

    # 세션 상태 초기화
    if "email_input" not in st.session_state:
        st.session_state.email_input = ""

    # 사이드바 - 설정
    with st.sidebar:
        st.header("🔧 설정")
        from dotenv import load_dotenv

        load_dotenv()

        # Azure OpenAI 설정
        st.subheader("Azure OpenAI 설정")

        # 환경변수에서 값 가져오기
        env_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
        env_api_key = os.getenv("AZURE_OPENAI_API_KEY")

        azure_endpoint = st.text_input(
            "Azure OpenAI Endpoint",
            value=env_endpoint if env_endpoint else "",
            placeholder="https://your-resource.openai.azure.com/",
            help="Azure OpenAI 리소스의 엔드포인트 URL. 환경변수 `AZURE_OPENAI_ENDPOINT`로 설정 가능",
        )

        api_key = st.text_input(
            "API Key",
            type="password",
            value=env_api_key if env_api_key else "",
            help="Azure OpenAI API 키. 환경변수 `AZURE_OPENAI_API_KEY`로 설정 가능",
        )

        model_name = st.text_input(
            "모델 이름",
            value="gpt-4",
            help="배포된 모델의 이름 (예: gpt-4, gpt-35-turbo)",
        )

        api_version = st.selectbox(
            "API 버전",
            ["2023-12-01-preview", "2023-10-01-preview", "2023-08-01-preview"],
            index=0,
        )

        st.markdown("---")

    # 메인 영역
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📥 E-mail Infomation")

        tab1, tab2 = st.tabs(["✍️ 텍스트로 입력/수정", "📎 파일로 업로드"])

        with tab2:
            # 파일 업로드
            uploaded_file = st.file_uploader(
                "이메일 파일 선택",
                type=["eml", "msg", "txt"],
                help="EML, MSG, 또는 텍스트 파일을 업로드하세요. 업로드 후 '텍스트로 입력/수정' 탭에서 내용을 확인하세요.",
            )

            if uploaded_file:
                try:
                    file_content = uploaded_file.read()
                    file_extension = uploaded_file.name.split(".")[-1].lower()

                    parsed_content = ""
                    if file_extension == "eml":
                        parsed_email = EmailParser.parse_eml_file(file_content)
                        if parsed_email:
                            parsed_content = parsed_email["full_content"]
                            st.success(
                                f"✅ EML 파일({uploaded_file.name})을 성공적으로 로드했습니다."
                            )
                        else:
                            st.error("EML 파일 파싱에 실패했습니다.")

                    elif file_extension == "msg":
                        parsed_email = EmailParser.parse_msg_file(file_content)
                        if parsed_email:
                            parsed_content = parsed_email["full_content"]
                            st.success(
                                f"✅ MSG 파일({uploaded_file.name})을 성공적으로 로드했습니다."
                            )
                        else:
                            st.error("MSG 파일 파싱에 실패했습니다.")

                    elif file_extension == "txt":
                        parsed_content = file_content.decode("utf-8", errors="ignore")
                        st.success(
                            f"✅ 텍스트 파일({uploaded_file.name})을 성공적으로 로드했습니다."
                        )

                    if parsed_content:
                        st.session_state.email_input = parsed_content
                        st.info(
                            "파일 내용이 로드되었습니다. '텍스트로 입력/수정' 탭에서 확인 및 수정할 수 있습니다."
                        )

                except Exception as e:
                    st.error(f"파일 처리 중 오류 발생: {str(e)}")

        with tab1:
            # 텍스트 입력
            st.session_state.email_input = st.text_area(
                "이메일 내용을 입력하거나 수정하세요:",
                value=st.session_state.email_input,
                height=400,
                placeholder="이곳에 이메일 내용을 붙여넣거나, 파일 업로드 후 내용을 확인/수정하세요.",
                help="분석할 이메일의 전체 내용을 입력하세요.",
            )

        # 분석 버튼
        analyze_button = st.button(
            "🔍 이메일 분석하기", type="primary", use_container_width=True
        )

    with col2:
        st.subheader("📊 Email Analysis Result")

        if analyze_button:
            if not azure_endpoint or not api_key:
                st.error("Azure OpenAI 설정을 완료해주세요.")
                return

            email_to_analyze = st.session_state.email_input
            if not email_to_analyze.strip():
                st.error("분석할 이메일 내용을 입력해주세요.")
                return

            # 분석 진행
            with st.spinner("이메일을 분석하는 중입니다... 잠시만 기다려주세요."):
                analyzer = EmailAnalyzer()

                if analyzer.initialize_azure_openai(
                    azure_endpoint, api_key, api_version
                ):
                    result = analyzer.analyze_email(email_to_analyze, model_name)

                    if result:
                        # 분석 결과 표시
                        display_analysis_result(result)

                        # 세션에 결과 저장
                        if "analysis_history" not in st.session_state:
                            st.session_state.analysis_history = []

                        st.session_state.analysis_history.append(
                            {
                                "timestamp": datetime.now(),
                                "email": (
                                    email_to_analyze[:100] + "..."
                                    if len(email_to_analyze) > 100
                                    else email_to_analyze
                                ),
                                "email_full": email_to_analyze,
                                "result": result,
                            }
                        )

                        # 내보내기 및 일정 연동 버튼들
                        st.markdown("---")
                        st.subheader("📤 결과 내보내기 & 일정 연동")

                        export_col1, export_col2, export_col3 = st.columns(3)

                        with export_col1:
                            # PDF 내보내기
                            if st.button(
                                "📄 PDF 내보내기",
                                use_container_width=True,
                                key="pdf_export",
                            ):
                                try:
                                    pdf_data = ExportManager.create_pdf_report(
                                        result, email_to_analyze
                                    )
                                    st.download_button(
                                        label="📥 PDF 다운로드",
                                        data=pdf_data,
                                        file_name=f"email_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True,
                                    )
                                except Exception as e:
                                    st.error(f"PDF 생성 오류: {str(e)}")

                        with export_col2:
                            # Excel 내보내기
                            if st.button(
                                "📊 Excel 내보내기",
                                use_container_width=True,
                                key="excel_export",
                            ):
                                try:
                                    excel_data = ExportManager.create_excel_report(
                                        result
                                    )
                                    st.download_button(
                                        label="📥 Excel 다운로드",
                                        data=excel_data,
                                        file_name=f"email_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True,
                                    )
                                except Exception as e:
                                    st.error(f"Excel 생성 오류: {str(e)}")

                        with export_col3:
                            # 일정 내보내기
                            if result.get("tasks") and st.button(
                                "📅 일정 내보내기",
                                use_container_width=True,
                                key="cal_export",
                            ):
                                try:
                                    email_subject = (
                                        email_to_analyze.split("\n")[0]
                                        .replace("제목:", "")
                                        .strip()
                                    )
                                    calendar_data = ExportManager.create_calendar_zip(
                                        result["tasks"], email_subject
                                    )
                                    st.download_button(
                                        label="📥 일정 다운로드 (ZIP)",
                                        data=calendar_data,
                                        file_name=f"calendar_tasks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                        mime="application/zip",
                                        use_container_width=True,
                                    )
                                except Exception as e:
                                    st.error(f"일정 생성 오류: {str(e)}")

                        # 일정 미리보기
                        if result.get("tasks"):
                            st.markdown("---")
                            st.subheader("📅 일정 미리보기")
                            calendar_df = CalendarIntegration.create_calendar_summary(
                                result["tasks"]
                            )
                            st.dataframe(
                                calendar_df, use_container_width=True, hide_index=True
                            )

                            # 개별 일정 파일 다운로드
                            st.markdown("**개별 일정 다운로드:**")
                            for i, task in enumerate(result["tasks"], 1):
                                col_task, col_download = st.columns([3, 1])

                                with col_task:
                                    st.write(f"{i}. {task.get('task', '')}")

                                with col_download:
                                    email_subject = (
                                        email_to_analyze.split("\n")[0]
                                        .replace("제목:", "")
                                        .strip()
                                    )
                                    ics_content = CalendarIntegration.create_ics_event(
                                        task, email_subject
                                    )
                                    if ics_content:
                                        st.download_button(
                                            label="📅",
                                            data=ics_content,
                                            file_name=f"task_{i}.ics",
                                            mime="text/calendar",
                                            key=f"download_task_{i}",
                                        )

        # 이전 분석 결과가 있다면 표시
        elif (
            "analysis_history" in st.session_state and st.session_state.analysis_history
        ):
            st.info(
                "👆 위의 '이메일 분석하기' 버튼을 클릭하여 새로운 분석을 시작하세요."
            )
            with st.expander("📜 최근 분석 결과 보기"):
                latest_result = st.session_state.analysis_history[-1]["result"]
                display_analysis_result(latest_result)

                # 최근 결과에 대한 내보내기 기능
                st.markdown("---")
                st.subheader("📤 최근 결과 내보내기")

                recent_col1, recent_col2, recent_col3 = st.columns(3)
                latest_email = st.session_state.analysis_history[-1]["email_full"]

                with recent_col1:
                    try:
                        pdf_data = ExportManager.create_pdf_report(
                            latest_result, latest_email
                        )
                        st.download_button(
                            label="📄 PDF 다운로드",
                            data=pdf_data,
                            file_name=f"recent_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf",
                            key="recent_pdf",
                        )
                    except Exception as e:
                        st.error(f"PDF 생성 오류: {str(e)}")

                with recent_col2:
                    try:
                        excel_data = ExportManager.create_excel_report(latest_result)
                        st.download_button(
                            label="📊 Excel 다운로드",
                            data=excel_data,
                            file_name=f"recent_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="recent_excel",
                        )
                    except Exception as e:
                        st.error(f"Excel 생성 오류: {str(e)}")

                with recent_col3:
                    if latest_result.get("tasks"):
                        try:
                            email_subject = (
                                latest_email.split("\n")[0].replace("제목:", "").strip()
                            )
                            calendar_data = ExportManager.create_calendar_zip(
                                latest_result["tasks"], email_subject
                            )
                            st.download_button(
                                label="📅 일정 다운로드 (ZIP)",
                                data=calendar_data,
                                file_name=f"recent_calendar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip",
                                key="recent_calendar",
                            )
                        except Exception as e:
                            st.error(f"일정 생성 오류: {str(e)}")

    # 분석 히스토리
    if (
        "analysis_history" in st.session_state
        and len(st.session_state.analysis_history) > 1
    ):
        st.markdown("---")
        st.subheader("📜 분석 히스토리")

        # 히스토리 관리 버튼
        hist_col1, hist_col2 = st.columns([1, 4])

        with hist_col1:
            if st.button("🗑️ 히스토리 삭제", help="모든 분석 히스토리를 삭제합니다"):
                st.session_state.analysis_history = []
                st.rerun()

        with hist_col2:
            # 전체 히스토리 Excel 내보내기
            if st.button("📊 전체 히스토리 Excel 내보내기"):
                try:
                    all_data = []
                    for i, history in enumerate(st.session_state.analysis_history):
                        result = history["result"]
                        # 기본 정보 추가
                        base_info = {
                            "분석번호": i + 1,
                            "분석일시": history["timestamp"].strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                            "이메일미리보기": str(history["email"])[:100],
                            "요약": str(result.get("summary", ""))[:200],
                            "긴급도": str(result.get("urgency_level", "")),
                            "감정톤": str(result.get("sentiment", "")),
                        }

                        # 할일이 있는 경우
                        if result.get("tasks") and len(result["tasks"]) > 0:
                            for j, task in enumerate(result["tasks"]):
                                row_data = base_info.copy()
                                row_data.update(
                                    {
                                        "할일번호": j + 1,
                                        "할일내용": str(task.get("task", ""))[:100],
                                        "우선순위": str(task.get("priority", "")),
                                        "마감일": (
                                            str(task.get("deadline", ""))
                                            if task.get("deadline")
                                            and task.get("deadline") != "null"
                                            else "미정"
                                        ),
                                        "담당자": (
                                            str(task.get("assignee", ""))
                                            if task.get("assignee")
                                            and task.get("assignee") != "null"
                                            else "미정"
                                        ),
                                    }
                                )
                                all_data.append(row_data)
                        else:
                            # 할일이 없는 경우에도 기본 정보는 추가
                            row_data = base_info.copy()
                            row_data.update(
                                {
                                    "할일번호": 0,
                                    "할일내용": "할일 없음",
                                    "우선순위": "-",
                                    "마감일": "-",
                                    "담당자": "-",
                                }
                            )
                            all_data.append(row_data)

                    if all_data:
                        df = pd.DataFrame(all_data)
                        buffer = io.BytesIO()

                        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                            df.to_excel(writer, sheet_name="전체히스토리", index=False)

                        buffer.seek(0)

                        st.download_button(
                            label="📥 전체 히스토리 다운로드",
                            data=buffer.getvalue(),
                            file_name=f"email_analysis_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="history_excel_download",
                        )
                    else:
                        st.warning("내보낼 히스토리 데이터가 없습니다.")
                except Exception as e:
                    st.error(f"히스토리 Excel 생성 오류: {str(e)}")
                    st.info("히스토리 내보내기에 실패했습니다. 데이터를 확인해주세요.")

        for i, history in enumerate(reversed(st.session_state.analysis_history[:-1])):
            with st.expander(
                f"분석 {len(st.session_state.analysis_history)-i-1}: {history['timestamp'].strftime('%Y-%m-%d %H:%M')}"
            ):
                st.text(f"이메일 미리보기: {history['email']}")
                display_analysis_result(history["result"])

                # 개별 히스토리 내보내기
                st.markdown("**이 결과 내보내기:**")
                hist_export_col1, hist_export_col2, hist_export_col3 = st.columns(3)

                with hist_export_col1:
                    try:
                        pdf_data = ExportManager.create_pdf_report(
                            history["result"], history["email_full"]
                        )
                        st.download_button(
                            label="📄 PDF",
                            data=pdf_data,
                            file_name=f"analysis_{i}_{history['timestamp'].strftime('%Y%m%d_%H%M%S')}.pdf",
                            mime="application/pdf",
                            key=f"hist_pdf_{i}",
                        )
                    except Exception as e:
                        st.error(f"PDF 생성 오류: {str(e)}")

                with hist_export_col2:
                    try:
                        excel_data = ExportManager.create_excel_report(
                            history["result"]
                        )
                        st.download_button(
                            label="📊 Excel",
                            data=excel_data,
                            file_name=f"analysis_{i}_{history['timestamp'].strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"hist_excel_{i}",
                        )
                    except Exception as e:
                        st.error(f"Excel 생성 오류: {str(e)}")

                with hist_export_col3:
                    if history["result"].get("tasks"):
                        try:
                            email_subject = (
                                history["email_full"]
                                .split("\n")[0]
                                .replace("제목:", "")
                                .strip()
                            )
                            calendar_data = ExportManager.create_calendar_zip(
                                history["result"]["tasks"], email_subject
                            )
                            st.download_button(
                                label="📅 일정",
                                data=calendar_data,
                                file_name=f"calendar_{i}_{history['timestamp'].strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip",
                                key=f"hist_calendar_{i}",
                            )
                        except Exception as e:
                            st.error(f"일정 생성 오류: {str(e)}")

    # 사용법 안내
    with st.expander("ℹ️ 사용법 안내"):
        st.markdown(
            """
        ### 📧 E-mail Analyzer 사용법
        
        #### 1. 설정
        - (`AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_KEY`) 환경변수 설정
        
        #### 2. 이메일 입력
        **방법 1: 텍스트로 입력/수정**
        - '텍스트로 입력/수정' 탭에서 이메일 내용을 직접 붙여넣거나 수정합니다.
        
        **방법 2: 파일로 업로드**
        - '파일로 업로드' 탭에서 EML, MSG, TXT 파일을 업로드합니다.
        - 업로드된 내용은 '텍스트로 입력/수정' 탭에 자동으로 표시됩니다.
                
        #### 3. 분석 결과 활용
        **화면 보기**
        - 요약, 할일 목록, 긴급도 등을 바로 확인합니다.
        
        **내보내기**
        - 📄 **PDF**: 완전한 보고서 형태로 내보냅니다.
        - 📊 **Excel**: 할일 목록을 표 형태로 관리합니다.
        - 📅 **일정**: ICS 파일로 캘린더 앱에 바로 추가합니다.
        
        #### 4. 히스토리 관리
        - 모든 분석 결과는 자동 저장됩니다.
        - 이전 분석 결과를 재확인하고 내보낼 수 있습니다.
        - 전체 히스토리를 Excel로 통합 관리할 수 있습니다.
        
        ### 💡 팁
        - 이메일이 길수록 더 정확한 분석이 가능합니다.
        - 마감일이 명시된 업무는 자동으로 일정에 반영됩니다.
        """
        )


def display_analysis_result(result: Dict[str, Any]):
    """분석 결과를 화면에 표시"""
    if not result:
        st.error("분석 결과를 가져올 수 없습니다.")
        return

    # 긴급도와 감정 지표
    col1, col2, col3 = st.columns(3)
    with col1:
        urgency_color = {"high": "🔴", "medium": "🟡", "low": "🟢"}
        st.metric(
            "긴급도",
            f"{urgency_color.get(result.get('urgency_level', 'medium'), '🟡')} {result.get('urgency_level', 'medium').upper()}",
        )

    with col2:
        sentiment_color = {"positive": "😊", "neutral": "😐", "negative": "😟"}
        st.metric(
            "감정 톤",
            f"{sentiment_color.get(result.get('sentiment', 'neutral'), '😐')} {result.get('sentiment', 'neutral').upper()}",
        )

    with col3:
        task_count = len(result.get("tasks", []))
        st.metric("할일 개수", f"📋 {task_count}개")

    # 요약
    st.subheader("📝 요약")
    st.info(result.get("summary", "요약 정보가 없습니다."))

    # 주요 포인트
    if result.get("key_points"):
        st.subheader("🎯 주요 포인트")
        for point in result["key_points"]:
            st.write(f"• {point}")

    # 할일 목록
    if result.get("tasks"):
        st.subheader("✅ 할일 목록")
        for i, task in enumerate(result["tasks"], 1):
            priority_colors = {"high": "🔴", "medium": "🟡", "low": "🟢"}

            priority_icon = priority_colors.get(task.get("priority", "medium"), "🟡")

            with st.container():
                st.write(f"**{i}. {task.get('task', '')}**")

                task_col1, task_col2, task_col3 = st.columns(3)
                with task_col1:
                    st.write(
                        f"우선순위: {priority_icon} {task.get('priority', 'medium')}"
                    )
                with task_col2:
                    deadline = task.get("deadline")
                    if deadline and deadline != "null":
                        st.write(f"마감일: 📅 {deadline}")
                    else:
                        st.write("마감일: 미정")
                with task_col3:
                    assignee = task.get("assignee")
                    if assignee and assignee != "null":
                        st.write(f"담당자: 👤 {assignee}")
                    else:
                        st.write("담당자: 미정")

                st.markdown("---")

    # 즉시 처리 항목
    if result.get("action_items"):
        st.subheader("⚡ 즉시 처리 항목")
        for item in result["action_items"]:
            st.error(f"🚨 {item}")

    # 후속 조치
    if result.get("follow_up"):
        st.subheader("🔄 후속 조치")
        st.warning(result["follow_up"])

    # Raw response (디버깅용)
    if result.get("raw_response"):
        with st.expander("🔧 원시 응답 (디버깅)"):
            st.text(result["raw_response"])


if __name__ == "__main__":
    main()
