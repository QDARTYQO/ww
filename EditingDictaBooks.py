from logging import info
import sys
import ctypes
from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel,
    QLayout, QFileDialog, QLineEdit, QMessageBox, QComboBox, QHBoxLayout,
    QCheckBox, QTextEdit, QDialog, QFrame, QSplitter, QGridLayout, QSpacerItem, QSizePolicy, QApplication
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QCursor, QColor, QPalette
from PyQt5.QtWinExtras import QtWin
from PyQt5.QtWidgets import QProxyStyle, QMessageBox, QTreeWidget
from PyQt5.QtCore import pyqtSignal, QThread, pyqtSignal, QTimer
from pyluach import gematria
from bs4 import BeautifulSoup
import re
import os
import requests
import sys
import shutil
from packaging import version
import base64
import urllib.request
import traceback



 #פונקצייה גלובלית לטיפול בשגיאות
def handle_exception(exc_type, exc_value, exc_traceback):
    """טיפול בשגיאות לא מטופלות"""
    print(''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
    
sys.excepthook = handle_exception


 #עיצוב גלובלי
GLOBAL_STYLE = """
    QWidget {
        font-family: "Segoe UI", Arial;
        font-size: 14px;
    }
    QPushButton {
        font-size: 20px;
    }
    QLabel {
        font-size: 20px;
    }
    QComboBox {
        font-size: 20px;
    }
    QLineEdit {
        font-size: 20px;
    }
    QCheckBox {
        font-size: 20px;
    }
"""

# מזהה ייחודי לאפליקציה
myappid = 'MIT.LEARN_PYQT.dictatootzaria'

# מחרוזת Base64 של האייקון (החלף את זה עם המחרוזת שתקבל אחרי המרת הקובץ שלך ל־Base64)
icon_base64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAGP0lEQVR4Ae2dfUgUaRjAn2r9wkzs3D4OsbhWjjKk7EuCLY9NrETJzfujQzgNBBEEIQS9/EOM5PCz68Ly44gsNOjYXVg/2NDo6o8iWg/ME7/WVtJ1XTm7BPUky9tnQLG7dnd2dnee9eb9waDs7DzPu/ObnX3nnZlnZMADpVIZLpfLUywWyzejo6MHbDZbtP3lED7LMpwjczYzNTX1q+np6R+ePn36HbAV7hMcCkhJSSnQ6/U/2v8NErE9kuM/Avbv3x8YEBDQ0N7e/j1Fg6TGJwJw5QcFBf1qNBpTqRokNT4RYN/yf2ErX1xWBSQnJxcaDIZMysZIEU6AWq3+WqPRXKVujBThBLx+/brE/ieAuC2SRHbq1Kkvurq6vqVuiFSRRUREpAHr65MhGx8fT6RuhJSRmUymA9SNkDIym832JWUD9u7diweAYD8AFC3nwsIC9PT0YOdDtJyOwF6Qy09eU1MDCoXC4fyqqip48uSJW4k3b94MN2/ehKSkJLeW8xbLy8tw//59KCwshKWlJUExXK2XsrIyePnypdMYTkdDVzhx4gQcOnTI4fzW1lY+YVbZtGkTt8yRI0fcWs6bbNiwAS5cuAAhISGQm5srKIar9dLQ0OAyBi8B3iYzM5N05a/l3Llz0NLS4vY32FuQCEhLS6NI6xCUICkBO3fupEjrkOjoaLLcJALwB9Cf2LhxI1luEgGOaGpqgvLycp/Fb2xsBJVK5bP4QvArAe/fv4f5+Xmfxf/w4YPPYgvFrwRIESaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGL8SsG/fPu5kja+IioryWWyh+JWAkydPcpOUYOcDgLY9JAIGBwedXk0gNkNDQ2S5SQTU19fDmTNnSM9ErYDnIG7fvk2Wn0TAixcvoKKiAoqKiijSr/Lx40e4fPmy9L4ByLVr12BsbAwKCgogJiZG1G8Drvj+/n6orKwEg8EgWt7PQdoL0mq13CSTySAwMFC0vIuLi35zetIvuqF4aaDQywPXO34hQMowAcQwAcQwAcQwAcSQCjh8+DAWBSHLj13g3t5esvwIqYCjR49CaWkpWX6z2SxtAQwmgBwmgBgmgBgmgBhSAXizNA4JU/Hq1Suy3CuQCnj+/Dk3SRm2CyKGCSCGCSCGCSCGCSCGVACWCNi9ezdZ/oGBAbDZbGT5EVIB2dnZpKOhWVlZcOfOHbL8CNsFEcMEEMMEEMMEEMMEEMMEEEMqYHZ2FiwWC1l+rB9KDamA2tpabpIybBdEDBNADC8Brm5mELPusz8RFhbmcQxeAt6+fet0fmxsrMcNWW8EBwfDrl27nL5nbm7OZRxeAkwmk9P5GRkZUFxcLKm7XPAzu/rmj4yMuIzDSwDe1ZiXl+dwPg4po4ArV67wCbfu2bp1K1y96vyZR1NTU/DmzRuXsXgJwDsJ8c5CZ3cy4rAy1uO/d+8en5DrFrlcDm1tbS7LHXd0dPCKx0uA1WrlLuU+f/68w/egnObmZq4kvE6n43ZbnuyS8AbqZ8+eCV5+LUqlkitX7wnh4eGQkJAAFy9ehB07djh9L5Y+uHXrFq+4vLuhuIVjlXGs/e8I/JBnz57lJk/BM1Xbt2/3OA7S2dkJoaGhXonFB71ez+22+cBbQF9fH9TV1UF+fr7ghkkBHN64dOkS7/e7dSCGpQVw696zZ4/bDZMKuPL59H5WcEsAFtbGgkqPHz+W7MGXM+7evcs9F8cd3B6KwGs58QkYDx48gC1btri7+P8W7PXk5OS4vZygsaCHDx9CXFwcVFdXQ3p6ul+UnaEC+/vYQcESPAIKP80IHozDSid4NBgfHw8lJSX4/Hmu6IZUmJiYgBs3bsD169cFP/PAfiwx7PEaw2v81Wo1bNu2jfuLNd/wR9rT3dPMzIynTVtleHiYe1yVJ+CA5OTkJBiNRq5biw/9wYNTT1AoFD1e22Sx344HH3wPQMTk4MGD1E34LJGRkY+ks8/wP2bNZnMnE0CESqX6ubu7e44JIMD+e/nHu3fvuMdFMQHi89fx48fTdTod13ViAsRl3t5TVGs0muGVF5gAkbB3g2dOnz6NK/+3ta8zASJg3+f/fuzYsQytVjv673lMgA+xb/V/JiYm/mS1Wiv1ev3fn3vPP+R95FTm9cojAAAAAElFTkSuQmCC="

 #מחלקה לטעינת קבצים
class FileLoader(QThread):
    """מחלקה לטעינת קבצים ברקע"""
    finished = pyqtSignal(dict)  # שולח מילון עם התוצאות

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        result = {
            'success': False,
            'content': None,
            'error': None
        }
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                result['success'] = True
                result['content'] = content
        except UnicodeDecodeError:
            result['error'] = "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8."
        except Exception as e:
            result['error'] = f"שגיאה בפתיחת הקובץ: {str(e)}"
        finally:
            self.finished.emit(result)

class DefaultTextContent:
    """מחלקה המכילה את טקסט ברירת המחדל לתצוגה ראשונית"""
    
    @staticmethod
    def get_default_html():
        return """<h1>עריכת ספר באוצריא</h1>
הסבר כיצד ניתן 'לערוך' את הטקסט כדי שיופיע בתוכנה כאילו הוא טקסט ערוך.
אנו צריכים להוסיף לטקסט לפני ואחרי המילה (או המילים) שאותם אנו רוצים לערוך, את הסימנים: < >, וביניהם אותיות מסוימות באנגלית [הנקראים תגי HTML, ותגי CSS], ובתג שאחרי המילה (או המילים) יש להוסיף גם קו נטוי בתוך התג לפני האותיות באנגלית [/].
התג הראשון מגדיר לתוכנה להתחיל להפעיל את צורת העריכה הכתובה בו, והתג האחרון מורה לתוכנה שעד כאן יש להפעיל זאת. הראשון נקרא תג פותח, והשני תג סוגר.
על מנת להבין איך יוצרים את כל סוגי האפשרויות האלו, יש לפתוח קובץ זה דרך קורא טקסטים כגון פנקס רשימות או וורד וכדומה וכך ניתן לראות את התגים.
מומלץ מאוד לאחר פתיחת קובץ זה ע"י עורך טקסט, ללחוץ על צירוף המקשים: Ctrl+Shift [קונטרול+שיפט] השמאליים במקלדת,  כדי ליישר את הטקסט לצד שמאל, כך ניתן להבין טוב יותר את צורת כתיבת התגים האלו. בכל עת ניתן להחזיר לצד ימין ע"י לחיצה על צירוף המקשים האלו שבצד ימין של המקלדת.
שימו לב! בכתיבת תג פותח בלי תג סוגר, התוכנה תמשיך את העריכה על כל הקטע עד ירידת השורה הבאה [ע"י Enter או Shift+Enter (אנטר או שיפט+אנטר)]. 
באם נכתב תג סוגר, הטקסט שנכתב מכאן ואילך גם אם הוא באותו קטע לא יוחל עליו ההגדרות הקודמות. יוצאים מכלל זה הם תגי הכותרות שמוחלות באופן אוטומטי על כל הפיסקה, גם אם נכתב באמצעה תג סוגר.
בכל פיסקה [ירידת שורה ע"י אנטר או שיפט+אנטר] יש לכתוב מחדש את התגים שברצונכם להפעיל בקטע זה.

<h1>כותרת רמה 1 [תמיד זה שם הספר, בשורה הראשונה]</h1>
<h2>כותרת רמה 2</h2>
<h3>כותרת רמה 3</h3>
<h4>כותרת רמה 4</h4>
<h5>כותרת רמה 5</h5>
<h6>כותרת רמה 6</h6>
שימו לב! כותרות ברמה 7 ומעלה, לא נתמכות.

<h2>הדגשה</h2>
<b>הדגשה</b>
<strong>צורה נוספת</strong>

<h2>גודל כתב</h2>
<big>כתב גדול</big> 
<big><big>כתב גדול מאוד</big></big>
<small>כתב קטן</small>
<small><small>כתב קטן מאוד</small></small>
[הסבר] להגדלת והקטנת גודל גופן מותאמת אישית יש לכתוב את התגים האלו לפני ואחרי הטקסט הרצוי.
<span style="font-size:150%;">הגדלת הכתב ל 150%</span>
<span style="font-size:70%;">הקטנת הכתב ל 70%</span>
<span style="font-size:50%;">הקטנת הכתב ל 50%</span>

<h2>כתב נטוי</h2>
<i>כתב נטוי</i>

<h2>סוג גופן</h2>
[הסבר] לבחירת גופן מסוים [ניתן לבחור גם גופן שלא נמצא ברשימת הגופנים בהגדרות התוכנה], יש לכתוב את התגים האלו לפני ואחרי הטקסט הרצוי.
<span style="font-family: SBL Hebrew;">כתב אחר</span>
<span style="font-family: Arial;">אריאל</span>
<span style="font-family: FrankRuehl;">פרנקריל</span>
<span style="font-family: Hadassah Friedlaender;">הדסה</span>
<span style="font-family: Narkisim;">נרקיסים</span>
<span style="font-family: Guttman Mantova;">מנטובה</span>
<span style="font-family: David;">דוד</span>
<span style="font-family: Guttman Logo1;">גוטמן</span>

<h2>צבע גופן</h2>
[הסבר] לבחירת צבע מסוים שבו יוצג הטקסט, יש להקליד את התגים האלו לפני ואחרי הטקסט הרצוי, כמובן ניתן לבחור כל סוג של צבע, יש לכתוב את שם הצבע באנגלית, לצבעים בסיסיים, או את קוד הצבע.
<span style="color:Black;">כאן לדוגמא נבחר הצבע השחור (ברירת מחדל, גם בלי לכתוב סוג צבע)</span>
<span style="color:Gray;">כאן לדוגמא נבחר הצבע האפור</span>
<span style="color:Blue;">כאן לדוגמא נבחר הצבע הכחול</span>
<span style="color:LightBlue;">כאן לדוגמא נבחר הצבע התכלת</span>
<span style="color:Green;">כאן לדוגמא נבחר הצבע הירוק</span>
<span style="color:LightGreen;">כאן לדוגמא נבחר הצבע הירוק בהיר</span>
<span style="color:Aqua;">כאן לדוגמא נבחר הצבע הטורקיז</span>
<span style="color:Red;">כאן לדוגמא נבחר הצבע האדום</span>
<span style="color:Purple;">כאן לדוגמא נבחר הצבע הסגול</span>
<span style="color:Yellow;">כאן לדוגמא נבחר הצבע הצהוב</span>
<span style="color:Lightyellow;">כאן לדוגמא נבחר הצבע הצהוב בהיר</span>
<span style="color:White;">כאן לדוגמא נבחר הצבע הלבן</span>
בחירת צבעים לפי קוד [ע"י קוד ניתן לבחור מגוון רחב מאוד של צבעים ותתי צבעים].
<span style="color:#00FF00;">כאן לדוגמא נבחר צבע ירוק זורח</span>
<span style="color:#2828AC;">כאן לדוגמא נבחר הצבע כחול כהה</span>
<span style="color:#F0E68C;">כאן לדוגמא נבחר הצבע הכתום</span>

<h2>יצירת מספר סוגי תגים ביחד</h2>
<span style="font-size:130%; font-family: SBL Hebrew; color:Red;">בחירת גופן והגדלת כתב ושינוי צבע ביחד</span>
[הסבר] אין הבדל בסדר הכתיבה מה רושמים קודם, את תגי סוג הגופן או הגודל או הצבע וכדו'.

<h2>הערות שוליים</h2>
בגוף הטקסט שבו רוצים לרשום את ההערה יש לרשום את התגים האלו<sup>2</sup> וביניהם את מספר ההערה<sup>1</sup>.
לאחר סוף הקטע יש לכתוב כך: 
<small><sup>1</sup> גוף ההערה</small>
<small><sup>2</sup> גוף ההערה</small>
כדי שמספר ההערה יכתב בצבע תכלת וכן עם קו מתחתיו [כמו קישור] נכתוב כך: [שים לב, קישורים עצמם אינם עובדים באוצריא, זה רק מראה של קישור].
הטקסט בגוף הספר<a href><sup>1</sup></a>.
<small><a href><sup>1</sup></a> גוף ההערה</small>
<small><a href><sup>2</sup></a> גוף ההערה</small>

<h2>סימנים נוספים</h2>
תו ר"ת הפוך &#8220;
תו חזור אחורה &#8617
תו חזור אחורה בצבע תכלת עם קו מתחתיו [כמו קישור] <a href>&#8617</a>"""

    @staticmethod
    def get_default_style():
        """מחזיר את הסגנון עבור תיבת הטקסט"""
        return """
            QTextEdit {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: white;
                border: 2px solid #2b4c7e;
                border-radius: 15px;
            }
        """           
# ==========================================
# Script 1: יצירת כותרות לאוצריא
# ==========================================

class CreateHeadersOtZria(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.setWindowTitle("יצירת כותרות לאוצריא")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 450)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)

        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        
        self.init_ui()

        if parent:
            parent_center = parent.mapToGlobal(parent.rect().center())
            self.move(parent_center.x() - self.width() // 2,
                     parent_center.y() - self.height() // 2)        

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש - בחלק העליון עם גוון אדום
        explanation = QLabel(
            "שים לב!\n\n"
            "בתיבת 'מילה לחפש' יש לבחור או להקליד את המילה בה אנו רוצים שתתחיל הכותרת.\n"
            "לדוג': פרק/פסוק/סימן/סעיף/הלכה/שאלה/עמוד/סק\n\n"
            "אין להקליד רווח אחרי המילה, וכן אין להקליד את התו גרש (') או גרשיים (\"), "
            "וכן אין להקליד יותר ממילה אחת"
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #8B0000;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;
                border: 2px solid #CD5C5C;
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # מילה לחיפוש
        search_container = QVBoxLayout()
        search_label = QLabel("מילה לחפש:")
        search_label.setStyleSheet(label_style)
        
        self.level_var = QComboBox()
        self.level_var.setStyleSheet(combo_style)
        self.level_var.setFixedWidth(150)
        search_choices = ["דף", "עמוד", "פרק", "פסוק", "שאלה", "סימן", "סעיף", "הלכה", "הלכות", "סק"]
        self.level_var.addItems(search_choices)
        self.level_var.setEditable(True)
        
        search_container.addWidget(search_label, alignment=Qt.AlignCenter)
        search_container.addWidget(self.level_var, alignment=Qt.AlignCenter)
        layout.addLayout(search_container)

        # מספר סימן מקסימלי
        end_container = QVBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        end_label.setStyleSheet(label_style)
        
        self.end_var = QComboBox()
        self.end_var.setStyleSheet(combo_style)
        self.end_var.setFixedWidth(100)
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        
        end_container.addWidget(end_label, alignment=Qt.AlignCenter)
        end_container.addWidget(self.end_var, alignment=Qt.AlignCenter)
        layout.addLayout(end_container)

        # רמת כותרת
        heading_container = QVBoxLayout()
        heading_label = QLabel("רמת כותרת:")
        heading_label.setStyleSheet(label_style)
        
        self.heading_level_var = QComboBox()
        self.heading_level_var.setStyleSheet(combo_style)
        self.heading_level_var.setFixedWidth(100)
        self.heading_level_var.addItems([str(i) for i in range(2, 7)])
        self.heading_level_var.setCurrentText("2")
        
        heading_container.addWidget(heading_label, alignment=Qt.AlignCenter)
        heading_container.addWidget(self.heading_level_var, alignment=Qt.AlignCenter)
        layout.addLayout(heading_container)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)

    def set_file_path(self, path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        self.file_path = path

    def show_custom_message(self, title, message_parts, window_size=("560x330")):
        msg = QMessageBox(self)
        msg.setStyleSheet(GLOBAL_STYLE)  # הוספת הסגנון הגלובלי
        msg.setWindowTitle(title)
        msg.setIcon(QMessageBox.Information)

        full_message = ""
        for part in message_parts:
            if len(part) == 3 and part[2] == "bold":
                full_message += f"<b><span style='font-size:{part[1]}pt'>{part[0]}</span></b><br>"
            else:
                full_message += f"<span style='font-size:{part[1]}pt'>{part[0]}</span><br>"

        msg.setTextFormat(Qt.RichText)
        msg.setText(full_message)
        msg.exec_()

    def ot(self, text, end):
        remove = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", ".", "״", "‚", "”", "’"]
        aa = ["ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק", "יה", "יו", "קיה", "קיו", "ריה", "ריו", "שיה", "שיו", "תיה", "תיו", "תקיה", "תקיו", "תריה", "תריו", "תשיה", "תשיו", "תתיה", "תתיו", "תתקיה", "תתקיו"]
        bb = ["ם", "ן", "ץ", "ף", "ך"]
        cc = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "ששי", "שביעי", "שמיני", "תשיעי", "עשירי", "יוד", "למד", "נון", "דש", "חי", "טל", "שדמ", "ער", "שדם", "תשדם", "תשדמ", "ערב", "ערה", "עדר", "רחצ"]
        append_list = []
        for i in aa:
            for ot_sofit in bb:
                append_list.append(i + ot_sofit)

        for tage in remove:
            text = text.replace(tage, "")
        withaute_gershayim = [gematria._num_to_str(i, thousands=False, withgershayim=False) for i in range(1, end)] + bb + cc + append_list
        return text in withaute_gershayim

    def strip_html_tags(self, text):
        html_tags = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", "״", ".", "‚", "”", "’"]
        for tag in html_tags:
            text = text.replace(tag, "")
        return text

    def main(self, book_file, finde, end, level_num):
        found = False
        count_headings = 0
        finde_cleaned = self.strip_html_tags(finde).strip()
        
        with open(book_file, "r", encoding="utf-8") as file_input:
            content = file_input.read().splitlines()
            all_lines = content[0:2]
            
            for line in content[2:]:
                words = line.split()
                try:
                    if self.strip_html_tags(words[0]) == finde and self.ot(words[1], end):
                        found = True
                        count_headings += 1
                        heading_line = f"<h{level_num}>{self.strip_html_tags(words[0])} {self.strip_html_tags(words[1])}</h{level_num}>"
                        all_lines.append(heading_line)
                        if words[2:]:
                            fix_2 = " ".join(words[2:])
                            all_lines.append(fix_2)
                    else:
                        all_lines.append(line)
                except IndexError:
                    all_lines.append(line)
                    
        join_lines = "\n".join(all_lines)
        with open(book_file, "w", encoding="utf-8") as autpoot:
            autpoot.write(join_lines)

        return found, count_headings
    
    def run_script(self):
        try:
            if not self.file_path:
                self.show_custom_message(
                    "שגיאה",
                    [("לא נבחר קובץ", 12)],
                    "250x80"
                )
                return

            finde = self.level_var.currentText()
            
            try:
                end = int(self.end_var.currentText())
                level_num = int(self.heading_level_var.currentText())
            except ValueError:
                self.show_custom_message(
                    "קלט לא תקין",
                    [("אנא הזן 'מספר סימן מקסימלי' ו'רמת כותרת' תקינים", 12)],
                    "250x150"
                )
                return

            if not finde:
                self.show_custom_message(
                    "קלט לא תקין",
                    [("אנא מלא את כל השדות", 12)],
                    "250x80"
                )
                return

            # הפעלת הפונקציה הראשית
            found, count_headings = self.main(self.file_path, finde, end + 1, level_num)

            # אם נבחרה המילה "דף" והיו שינויים, הפעל את סקריפט 3
            if finde == "דף" and found and count_headings > 0:
                add_page_number = AddPageNumberToHeading()
                add_page_number.set_file_path(self.file_path)
                add_page_number.process_file(self.file_path, "נקודה ונקודותיים")
                
                detailed_message = [
                    ("<div style='text-align: center;'>התוכנה רצה בהצלחה!</div>", 12),
                    (f"<div style='text-align: center;'>נוצרו {count_headings} כותרות והוספו מספרי עמודים</div>", 15, "bold"),
                    ("<div style='text-align: center;'>כעת פתח את הספר בתוכנת 'אוצריא', והשינויים ישתקפו ב'ניווט' שבתפריט הצידי.</div>", 11)
                ]
            elif found and count_headings > 0:
                detailed_message = [
                    ("<div style='text-align: center;'>התוכנה רצה בהצלחה!</div>", 12),
                    (f"<div style='text-align: center;'>נוצרו {count_headings} כותרות</div>", 15, "bold"),
                    ("<div style='text-align: center;'>כעת פתח את הספר בתוכנת 'אוצריא', והשינויים ישתקפו ב'ניווט' שבתפריט הצידי.</div>", 11)
                ]
                
            if found and count_headings > 0:
                self.show_custom_message("!מזל טוב", detailed_message, "560x310")
                self.changes_made.emit()
            else:
                self.show_custom_message("!שים לב", [("לא נמצא מה להחליף", 12)], "250x80")

        except Exception as e:
            self.show_custom_message("שגיאה", [("אירעה שגיאה: " + str(e), 12)], "250x150")

    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 2: יצירת כותרות לאותיות בודדות
# ==========================================

class CreateSingleLetterHeaders(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("יצירת כותרות לאותיות בודדות")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 600)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)

        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש בחלק העליון
        explanation = QLabel(
            "שים לב!\n\n"
            "הבחירה בברירת מחדל [השורה הריקה], משמעותה סימון כל האפשרויות."
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #8B0000;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;
                border: 2px solid #CD5C5C;
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # עיצוב משותף לתיבות טקסט
        entry_style = """
            QLineEdit {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                background-color: white;
            }
        """

        # תו בתחילת וסוף האות + רמת כותרת באותה שורה
        chars_container = QHBoxLayout()
        
        # קונטיינר לתו התחלה
        start_container = QVBoxLayout()
        start_label = QLabel("תו בתחילת האות:")
        start_label.setStyleSheet(label_style)
        self.start_var = QComboBox()
        self.start_var.addItems(["", "(", "["])
        self.start_var.setStyleSheet(combo_style)
        self.start_var.setFixedWidth(100)
        start_container.addWidget(start_label, alignment=Qt.AlignCenter)
        start_container.addWidget(self.start_var, alignment=Qt.AlignCenter)
        
        # קונטיינר לתו סוף
        end_container = QVBoxLayout()
        end_label = QLabel("תו/ים בסוף האות:")
        end_label.setStyleSheet(label_style)
        self.finde_var = QComboBox()
        self.finde_var.addItems(['', '.', ',', "'", "',", "'.", ']', ')', "']", "')", "].", ").", "],", "),", "'),", "').", "'],", "']."])
        self.finde_var.setStyleSheet(combo_style)
        self.finde_var.setFixedWidth(100)
        end_container.addWidget(end_label, alignment=Qt.AlignCenter)
        end_container.addWidget(self.finde_var, alignment=Qt.AlignCenter)
        
        # קונטיינר לרמת כותרת
        heading_container = QVBoxLayout()
        heading_label = QLabel("רמת כותרת:")
        heading_label.setStyleSheet(label_style)
        self.level_var = QComboBox()
        self.level_var.setStyleSheet(combo_style)
        self.level_var.setFixedWidth(100)
        self.level_var.addItems([str(i) for i in range(2, 7)])
        self.level_var.setCurrentText("3")
        heading_container.addWidget(heading_label, alignment=Qt.AlignCenter)
        heading_container.addWidget(self.level_var, alignment=Qt.AlignCenter)

        # הוספת מרווחים גמישים בין האלמנטים
        chars_container.addStretch(1)
        chars_container.addLayout(start_container)
        chars_container.addStretch(1)
        chars_container.addLayout(end_container)
        chars_container.addStretch(1)
        chars_container.addLayout(heading_container)
        chars_container.addStretch(1)
        
        layout.addLayout(chars_container)

        # תיבת סימון לחיפוש עם תווי הדגשה
        self.bold_var = QCheckBox("לחפש עם תווי הדגשה בלבד")
        self.bold_var.setStyleSheet("""
            QCheckBox {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
        """)
        self.bold_var.setChecked(True)
        layout.addWidget(self.bold_var, alignment=Qt.AlignCenter)

        # התעלם מהתווים
        ignore_container = QVBoxLayout()
        ignore_label = QLabel("התעלם מהתווים הבאים:")
        ignore_label.setStyleSheet(label_style)
        
        self.ignore_entry = QLineEdit()
        self.ignore_entry.setStyleSheet(entry_style)
        self.ignore_entry.setText('<big> </big> " ')
        
        ignore_container.addWidget(ignore_label, alignment=Qt.AlignCenter)
        ignore_container.addWidget(self.ignore_entry)
        layout.addLayout(ignore_container)

        # הסרת תווים
        remove_container = QVBoxLayout()
        remove_label = QLabel("הסר את התווים הבאים:")
        remove_label.setStyleSheet(label_style)
        
        self.remove_entry = QLineEdit()
        self.remove_entry.setStyleSheet(entry_style)
        self.remove_entry.setText('<b> </b> <big> </big> , : " \' . ( ) [ ] { }')
        
        remove_container.addWidget(remove_label, alignment=Qt.AlignCenter)
        remove_container.addWidget(self.remove_entry)
        layout.addLayout(remove_container)

        # מספר סימן מקסימלי
        end_container = QVBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        end_label.setStyleSheet(label_style)
        
        self.end_var = QComboBox()
        self.end_var.setStyleSheet(combo_style)
        self.end_var.setFixedWidth(100)
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        
        end_container.addWidget(end_label, alignment=Qt.AlignCenter)
        end_container.addWidget(self.end_var, alignment=Qt.AlignCenter)
        layout.addLayout(end_container)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)


        
    def set_file_path(self, file_path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.critical(self, "שגיאה", "נתיב קובץ לא תקין")
            return False
        
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "שגיאה", "יש לבחור קובץ טקסט (txt) בלבד")
            return False
        
        self.file_path = file_path
        return True

    def run_script(self):
        if not self.current_file_path:
            self.show_error_message("שגיאה", "אנא בחר קובץ תחילה")
            return
        
        finde = self.finde_var.currentText()
        remove = ["<b>", "</b>"] + self.remove_entry.text().split()
        ignore = self.ignore_entry.text().split()
        start = self.start_var.currentText()
        is_bold_checked = self.bold_var.isChecked()

        if is_bold_checked:
            finde += "</b>"
            start = "<b>" + start
        else:
            ignore += ["<b>", "</b>"]

        try:
            end = int(self.end_var.currentText())
            level_num = int(self.level_var.currentText())
        except ValueError:
            QMessageBox.critical(self, "קלט לא תקין", "אנא הזן 'מספר סימן מקסימלי' ו'רמת כותרת' תקינים")
            return

        try:
            self.main(self.file_path, finde, end + 1, level_num, ignore, start, remove)
            QMessageBox.information(self, "!מזל טוב", "התוכנה רצה בהצלחה!")
            self.changes_made.emit()  # שליחת סיגנל על שינויים
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה: {str(e)}")

    def ot(self, text, end):
        remove = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", ".", "״", "‚", "”", "’"]
        aa = ["ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק", "יו", "קיה", "קיו"]
        bb = ["ם", "ן", "ץ", "ף", "ך"]
        cc = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "ששי", "שביעי", "שמיני", "תשיעי", "עשירי", "חי", "יוד", "למד", "נון", "טל", "דש", "שדמ", "ער", "שדם", "תשדם", "תשדמ", "ערה", "ערב", "עדר", "רחצ"]
        append_list = []
        for i in aa:
            for ot_sofit in bb:
                append_list.append(i + ot_sofit)

        for tage in remove:
            text = text.replace(tage, "")
        withaute_gershayim = [gematria._num_to_str(i, thousands=False, withgershayim=False) for i in range(1, end)] + bb + cc + append_list
        return text in withaute_gershayim

    def strip_html_tags(self, text, ignore=None):
        if ignore is None:
            ignore = []
        for tag in ignore:
            text = text.replace(tag, "")
        return text

    def main(self, book_file, finde, end, level_num, ignore, start, remove):
        with open(book_file, "r", encoding="utf-8") as file_input:
            content = file_input.read().splitlines()
            all_lines = content[0:1]
            for line in content[1:]:
                words = line.split()
                try:
                    if self.strip_html_tags(words[0], ignore).endswith(finde) and self.ot(words[0], end) and self.strip_html_tags(words[0], ignore).startswith(start):
                        heading_line = f"<h{level_num}>{self.strip_html_tags(words[0], remove)}</h{level_num}>"
                        all_lines.append(heading_line)
                        if words[1:]:
                            fix_2 = " ".join(words[1:])
                            all_lines.append(fix_2)
                    else:
                        all_lines.append(line)
                except IndexError:
                    all_lines.append(line)
        join_lines = "\n".join(all_lines)
        with open(book_file, "w", encoding="utf-8") as autpoot:
            autpoot.write(join_lines)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   

# ==========================================
# Script 3: הוספת מספר עמוד בכותרת הדף משומש על ידי סריפט 1
# ==========================================
class AddPageNumberToHeading(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("הוספת מספר עמוד בכותרת הדף")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 600, 500)  # גודל חלון גדול יותר
        self.setLayoutDirection(Qt.RightToLeft)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)  # מרווחים מהשוליים

        # הסבר למשתמש
        explanation = QLabel(
            "התוכנה מחליפה בקובץ בכל מקום שיש כותרת 'דף' ובתחילת שורה הבאה כתוב: ע\"א או ע\"ב, כגון:\n\n"
            "<h2>דף ב</h2>\n"
            "ע\"א [טקסט כלשהו]\n\n"
            "הפעלת התוכנה תעדכן את הכותרת ל:\n\n"
            "<h2>דף ב.</h2>\n"
            "[טקסט כלשהו]\n"
        )
        explanation.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 2px solid black;
                border-radius: 15px;
                padding: 20px;
                font-family: "David CLM", Arial;
                font-size: 14px;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # מרווח קטן
        layout.addSpacing(15)

        # תיבת בחירה
        self.replace_option = QComboBox()
        self.replace_option.addItems(["נקודה ונקודותיים", "ע\"א וע\"ב"])
        self.replace_option.setStyleSheet("""
            QComboBox {
                border: 2px solid black;
                border-radius: 15px;
                padding: 5px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
            }
        """)
        self.replace_option.setFixedWidth(200)
        
        # מיכל למרכוז תיבת הבחירה
        combo_container = QHBoxLayout()
        combo_container.addStretch()
        combo_container.addWidget(self.replace_option)
        combo_container.addStretch()
        layout.addLayout(combo_container)

        # מרווח קטן
        layout.addSpacing(15)

        # כפתור הפעלה
        run_button = QPushButton("בצע החלפה")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                background-color: #eaeaea;
                border-radius: 15px;
                padding: 5px;
                font-family: "Segoe UI", Arial;
                font-weight: bold;
                font-size: 12px;
                min-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
        """)
        
        # מיכל למרכוז הכפתור
        button_container = QHBoxLayout()
        button_container.addStretch()
        button_container.addWidget(run_button)
        button_container.addStretch()
        layout.addLayout(button_container)

        # מרווח גמיש בסוף
        layout.addStretch()

        self.setLayout(layout)



    def set_file_path(self, path):
        self.file_path = path

    def process_file(self, filename, replace_with):
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                content = file.readlines()
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בפתיחת הקובץ: {str(e)}")
            return

        changes_made = False
        updated_content = []
        i = 0

        while i < len(content):
            line = content[i]
            match = re.match(r'<h([2-9])>(דף \S+)</h\1>', line)
            
            if match and i + 1 < len(content):
                level = match.group(1)
                title = match.group(2)
                next_line = content[i + 1].strip()
                
                if re.match(r'ע["\']א|ע["\']ב', next_line):
                    changes_made = True
                    if replace_with == "נקודה ונקודותיים":
                        suffix = "." if "א" in next_line else ":"
                    else:
                        suffix = " ע\"א" if "א" in next_line else " ע\"ב"
                    
                    updated_line = f"<h{level}>{title}{suffix}</h{level}>\n"
                    updated_content.append(updated_line)
                    
                    remaining_text = re.sub(r'^ע["\']א|ע["\']ב\s*', '', next_line)
                    if remaining_text:
                        updated_content.append(remaining_text + "\n")
                    i += 2
                    continue
            
            updated_content.append(line)
            i += 1

        if changes_made:
            try:
                with open(filename, 'w', encoding='utf-8') as file:
                    file.writelines(updated_content)
                self.changes_made.emit()
                QMessageBox.information(self, "הצלחה", "ההחלפות בוצעו בהצלחה!")
            except Exception as e:
                QMessageBox.critical(self, "שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")
        else:
            QMessageBox.information(self, "מידע", "לא נמצאו החלפות לביצוע")

    def run_script(self):
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        
        replace_with = self.replace_option.currentText()
        self.process_file(self.file_path, replace_with)

    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 3: שינוי רמת כותרת(4 לשעבר)
# ==========================================
class ChangeHeadingLevel(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("שינוי רמת כותרת")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 400)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש - עכשיו בחלק העליון עם גוון אדום
        explanation = QLabel(
            "שים לב!\n"
            "הכותרות יוחלפו מרמה נוכחית לרמה החדשה.\n"
            "למשל: מ-H2 ל-H3"
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #8B0000;  /* צבע טקסט אדום כהה */
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;  /* רקע אדום בהיר */
                border: 2px solid #CD5C5C;  /* מסגרת אדומה */
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 70px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # רמת כותרת נוכחית - עכשיו בסידור אנכי
        current_level_container = QVBoxLayout()
        current_level_label = QLabel("רמת כותרת נוכחית:")
        current_level_label.setStyleSheet(label_style)
        
        self.current_level_var = QComboBox()
        self.current_level_var.addItems([str(i) for i in range(1, 10)])
        self.current_level_var.setCurrentText("2")
        self.current_level_var.setStyleSheet(combo_style)
        
        current_level_container.addWidget(current_level_label, alignment=Qt.AlignCenter)
        current_level_container.addWidget(self.current_level_var, alignment=Qt.AlignCenter)
        layout.addLayout(current_level_container)

        # רמת כותרת חדשה - עכשיו בסידור אנכי
        new_level_container = QVBoxLayout()
        new_level_label = QLabel("רמת כותרת חדשה:")
        new_level_label.setStyleSheet(label_style)
        
        self.new_level_var = QComboBox()
        self.new_level_var.addItems([str(i) for i in range(1, 10)])
        self.new_level_var.setCurrentText("3")
        self.new_level_var.setStyleSheet(combo_style)
        
        new_level_container.addWidget(new_level_label, alignment=Qt.AlignCenter)
        new_level_container.addWidget(self.new_level_var, alignment=Qt.AlignCenter)
        layout.addLayout(new_level_container)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("שנה רמת כותרת")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)

    def set_file_path(self, file_path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.critical(self, "שגיאה", "נתיב קובץ לא תקין")
            return False
        
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "שגיאה", "יש לבחור קובץ טקסט (txt) בלבד")
            return False
        
        self.file_path = file_path
        return True

    def run_script(self):
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return

        current_level = self.current_level_var.currentText()
        new_level = self.new_level_var.currentText()

        try:
            changes_count = self.change_heading_level_func(
                self.file_path, 
                int(current_level), 
                int(new_level)
            )
            
            if changes_count > 0:
                QMessageBox.information(
                    self, 
                    "!מזל טוב", 
                    f"בוצעו {changes_count} החלפות בהצלחה!"
                )
                self.changes_made.emit()  # שליחת סיגנל על שינויים
            else:
                QMessageBox.information(
                    self, 
                    "!שים לב", 
                    "לא נמצאו כותרות להחלפה"
                )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "שגיאה", 
                f"אירעה שגיאה: {str(e)}"
            )

    def change_heading_level_func(self, file_path, current_level, new_level):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # יצירת דפוס חיפוש דינמי
            current_tag = f"h{current_level}"
            new_tag = f"h{new_level}"
            
            # ביטוי רגולרי להחלפת תגי כותרות
            pattern = re.compile(
                rf'<{current_tag}>(.*?)</{current_tag}>', 
                re.DOTALL | re.IGNORECASE
            )
            
            # ביצוע ההחלפה
            updated_content, changes_count = pattern.subn(
                lambda match: f'<{new_tag}>{match.group(1)}</{new_tag}>', 
                content
            )

            # אם בוצעו שינויים, שמירת הקובץ
            if changes_count > 0:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(updated_content)

            return changes_count

        except FileNotFoundError:
            QMessageBox.critical(self, "שגיאה", "הקובץ לא נמצא")
            return 0
        except UnicodeDecodeError:
            QMessageBox.critical(self, "שגיאה", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return 0
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")
            return 0

    def load_icon_from_base64(self, base64_string):
        """טעינת אייקון ממחרוזת Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
# ==========================================
# Script 4: הדגשת מילה ראשונה וניקוד בסוף קטע (5 לשעבר)
# ==========================================
class EmphasizeAndPunctuate(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("הדגשה וניקוד")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 400)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש בחלק העליון - בירוק בהיר
        explanation = QLabel(
            "הסבר:\n\n"
            "• הדגשת תחילת קטעים: מדגיש את המילה הראשונה בקטעים\n"
            "• הוספת סימן סוף: מוסיף נקודה או נקודותיים בסוף קטעים ארוכים"
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #1e4620;  /* ירוק כהה לטקסט */
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #e8f5e9;  /* ירוק בהיר לרקע */
                border: 2px solid #81c784;  /* ירוק בינוני למסגרת */
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # בחירה להוספת נקודה או נקודותיים
        ending_container = QVBoxLayout()
        ending_label = QLabel("בחר פעולה לסוף קטע:")
        ending_label.setStyleSheet(label_style)
        
        self.ending_var = QComboBox()
        self.ending_var.addItems(["הוסף נקודותיים", "הוסף נקודה", "ללא שינוי"])
        self.ending_var.setStyleSheet(combo_style)
        self.ending_var.setFixedWidth(170)
        
        ending_container.addWidget(ending_label, alignment=Qt.AlignCenter)
        ending_container.addWidget(self.ending_var, alignment=Qt.AlignCenter)
        layout.addLayout(ending_container)

        # הדגשת תחילת קטע
        self.emphasize_var = QCheckBox("הדגש את תחילת הקטעים")
        self.emphasize_var.setStyleSheet("""
            QCheckBox {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
        """)
        self.emphasize_var.setChecked(True)
        layout.addWidget(self.emphasize_var, alignment=Qt.AlignCenter)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)
        
    def set_file_path(self, file_path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.critical(self, "שגיאה", "נתיב קובץ לא תקין")
            return False
        
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "שגיאה", "יש לבחור קובץ טקסט (txt) בלבד")
            return False
        
        self.file_path = file_path
        return True

    def run_script(self):
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return

        try:
            changes_count = self.process_file(
                self.file_path, 
                self.ending_var.currentText(), 
                self.emphasize_var.isChecked()
            )
            
            if changes_count > 0:
                QMessageBox.information(
                    self, 
                    "!מזל טוב", 
                    f"בוצעו {changes_count} שינויים בהצלחה!"
                )
                self.changes_made.emit()  # שליחת סיגנל על שינויים
            else:
                QMessageBox.information(
                    self, 
                    "!שים לב", 
                    "לא נמצאו שינויים מתאימים בקובץ"
                )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "שגיאה", 
                f"אירעה שגיאה בעיבוד הקובץ: {str(e)}"
            )

    def process_file(self, file_path, add_ending, emphasize_start):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            changes_count = 0
            new_lines = []

            for line in lines:
                line = line.rstrip()
                words = line.split()

                # בדיקה אם יש יותר מעשר מילים ושאין סימן כותרת בהתחלה
                if (len(words) > 10 and 
                    not any(line.startswith(f'<h{n}>') for n in range(2, 10))):
                    
                    # הסרת רווחים ותווים מיותרים בסוף השורה
                    line = line.rstrip(" .,;:!?)</small></big></b>")

                    # הוספת סימן סוף
                    if add_ending != "ללא שינוי":
                        if line.endswith(','):
                            line = line[:-1]
                            line += '.' if add_ending == "הוסף נקודה" else ':'
                            changes_count += 1
                        elif not line.endswith(('.', ':', '!', '?')) and \
                             not any(line.endswith(tag) for tag in ['</small>', '</big>', '</b>']):
                            line += '.' if add_ending == "הוסף נקודה" else ':'
                            changes_count += 1

                    # הדגשת המילה הראשונה
                    if emphasize_start:
                        first_word = words[0]
                        if not any(tag in first_word for tag in ['<b>', '<small>', '<big>', '<h2>', '<h3>', '<h4>', '<h5>', '<h6>']):
                            if not (first_word.startswith('<') and first_word.endswith('>')):
                                line = f'<b>{first_word}</b> ' + ' '.join(words[1:])
                                changes_count += 1

                new_lines.append(line + '\n')

            # שמירת השינויים
            if changes_count > 0:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.writelines(new_lines)

            return changes_count

        except FileNotFoundError:
            QMessageBox.critical(self, "שגיאה", "הקובץ לא נמצא")
            return 0
        except UnicodeDecodeError:
            QMessageBox.critical(self, "שגיאה", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return 0
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")
            return 0

    def load_icon_from_base64(self, base64_string):
        """טעינת אייקון ממחרוזת Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
# ==========================================
# Script 5: יצירת כותרות לעמוד ב (6 לשעבר)
# ==========================================
class CreatePageBHeaders(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("יצירת כותרות עמוד ב")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 400)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש בחלק העליון - בירוק בהיר
        explanation = QLabel(
            "הסבר:\n\n"
            "• התוכנה תוסיף כותרת 'עמוד ב' לפני קטעים ללא כותרת\n"
            "• ניתן לבחור סוג כותרת ורמת כותרת שונים"
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #1e4620;  /* ירוק כהה לטקסט */
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #e8f5e9;  /* ירוק בהיר לרקע */
                border: 2px solid #81c784;  /* ירוק בינוני למסגרת */
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # סוג כותרת ורמת כותרת באותה שורה
        headers_container = QHBoxLayout()

        # קונטיינר לסוג כותרת
        header_type_container = QVBoxLayout()
        header_type_label = QLabel("סוג כותרת:")
        header_type_label.setStyleSheet(label_style)
        
        self.header_type_var = QComboBox()
        self.header_type_var.addItems([
            "עמוד ב", 
            "עמוד ב ע\"א", 
            "עמוד ב ע\"ב", 
            "עמוד ב'", 
            "עמוד ב׳"
        ])
        self.header_type_var.setStyleSheet(combo_style)
        self.header_type_var.setFixedWidth(170)
        
        header_type_container.addWidget(header_type_label, alignment=Qt.AlignCenter)
        header_type_container.addWidget(self.header_type_var, alignment=Qt.AlignCenter)

        # קונטיינר לרמת כותרת
        level_container = QVBoxLayout()
        level_label = QLabel("רמת כותרת:")
        level_label.setStyleSheet(label_style)
        
        self.level_var = QComboBox()
        self.level_var.addItems([str(i) for i in range(2, 7)])
        self.level_var.setCurrentText("3")
        self.level_var.setStyleSheet(combo_style)
        self.level_var.setFixedWidth(100)
        
        level_container.addWidget(level_label, alignment=Qt.AlignCenter)
        level_container.addWidget(self.level_var, alignment=Qt.AlignCenter)

        # הוספת מרווחים גמישים בין האלמנטים
        headers_container.addStretch(1)
        headers_container.addLayout(header_type_container)
        headers_container.addStretch(1)
        headers_container.addLayout(level_container)
        headers_container.addStretch(1)
        
        layout.addLayout(headers_container)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)


    def set_file_path(self, file_path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.critical(self, "שגיאה", "נתיב קובץ לא תקין")
            return False
        
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "שגיאה", "יש לבחור קובץ טקסט (txt) בלבד")
            return False
        
        self.file_path = file_path
        return True

    def run_script(self):
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return

        try:
            changes_count = self.process_file(
                self.file_path, 
                self.header_type_var.currentText(), 
                int(self.level_var.currentText())
            )
            
            if changes_count > 0:
                QMessageBox.information(
                    self, 
                    "!מזל טוב", 
                    f"בוצעו {changes_count} שינויים בהצלחה!"
                )
                self.changes_made.emit()  # שליחת סיגנל על שינויים
            else:
                QMessageBox.information(
                    self, 
                    "!שים לב", 
                    "לא נמצאו שינויים מתאימים בקובץ"
                )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "שגיאה", 
                f"אירעה שגיאה בעיבוד הקובץ: {str(e)}"
            )

    def process_file(self, file_path, header_type, heading_level):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            new_lines = []
            changes_count = 0
            is_first_paragraph = True

            for line in lines:
                # בדיקה אם השורה היא כותרת
                if any(line.startswith(f'<h{n}>') for n in range(2, 10)):
                    is_first_paragraph = False
                    new_lines.append(line)
                    continue

                # בדיקה אם השורה ריקה
                if not line.strip():
                    new_lines.append(line)
                    continue

                # אם זהו הקטע הראשון ללא כותרת
                if is_first_paragraph:
                    # הוספת כותרת עמוד ב
                    header_line = f'<h{heading_level}>{header_type}</h{heading_level}>\n'
                    new_lines.append(header_line)
                    changes_count += 1
                    is_first_paragraph = False

                new_lines.append(line)

            # שמירת השינויים
            if changes_count > 0:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.writelines(new_lines)

            return changes_count

        except FileNotFoundError:
            QMessageBox.critical(self, "שגיאה", "הקובץ לא נמצא")
            return 0
        except UnicodeDecodeError:
            QMessageBox.critical(self, "שגיאה", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return 0
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")
            return 0

    def load_icon_from_base64(self, base64_string):
        """טעינת אייקון ממחרוזת Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
# ==========================================
# Script 6: החלפת כותרות לעמוד ב (7 לשעבר)
# ==========================================
class ReplacePageBHeaders(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("החלפת כותרות ל'עמוד ב'")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 500, 500)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר ראשון - שים לב (באדום)
        attention = QLabel(
            "שים לב!\n\n"
            "התוכנה פועלת רק אם הדפים והעמודים הוגדרו כבר ככותרות\n"
            "[לא משנה באיזו רמת כותרת]\n"
            "כגון:  <h3>עמוד ב</h3> או: <h2>עמוד ב</h2> וכן הלאה"
        )
        attention.setStyleSheet("""
            QLabel {
                color: #8B0000;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;
                border: 2px solid #CD5C5C;
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        attention.setAlignment(Qt.AlignCenter)
        attention.setWordWrap(True)
        layout.addWidget(attention)

        # הסבר שני - זהירות (באדום)
        warning = QLabel(
            "זהירות!\n\n"
            "בדוק היטב שלא פספסת שום כותרת של 'דף' לפני שאתה מריץ תוכנה זו\n"
            "כי במקרה של פספוס, הכותרת 'עמוד ב' שאחרי הפספוס תהפך לכותרת שגויה"
        )
        warning.setStyleSheet("""
            QLabel {
                color: #8B0000;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;
                border: 2px solid #CD5C5C;
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        warning.setAlignment(Qt.AlignCenter)
        warning.setWordWrap(True)
        layout.addWidget(warning)

        # עיצוב משותף לתוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin-bottom: 5px;
            }
        """

        # עיצוב משותף לקומבו-בוקסים
        combo_style = """
            QComboBox {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px 15px;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                background-color: white;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #2b4c7e;
                margin-right: 5px;
            }
        """

        # סוג ההחלפה
        replace_container = QVBoxLayout()
        replace_label = QLabel("בחר את סוג ההחלפה:")
        replace_label.setStyleSheet(label_style)
        
        self.replace_type = QComboBox()
        self.replace_type.addItems(["נקודותיים", "ע\"ב"])
        self.replace_type.setStyleSheet(combo_style)
        self.replace_type.setFixedWidth(140)
        
        replace_container.addWidget(replace_label, alignment=Qt.AlignCenter)
        replace_container.addWidget(self.replace_type, alignment=Qt.AlignCenter)
        layout.addLayout(replace_container)

        # דוגמאות
        example1 = QLabel("לדוגמא:\nדף ב:   דף ג:   דף ד:   דף ה:\nוכן הלאה")
        example1.setStyleSheet("""
            QLabel {
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                color: #666;
            }
        """)
        example1.setAlignment(Qt.AlignCenter)
        layout.addWidget(example1)

        example2 = QLabel("או:\nדף ב ע\"ב   דף ג ע\"ב   דף ד ע\"ב   דף ה ע\"ב\nוכן הלאה")
        example2.setStyleSheet("""
            QLabel {
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                color: #666;
            }
        """)
        example2.setAlignment(Qt.AlignCenter)
        layout.addWidget(example2)

        # כפתור הפעלה
        button_container = QHBoxLayout()
        run_button = QPushButton("בצע החלפה")
        run_button.clicked.connect(self.run_script)
        run_button.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        button_container.addStretch(1)
        button_container.addWidget(run_button)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)


    def set_file_path(self, file_path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        if not file_path or not os.path.isfile(file_path):
            QMessageBox.critical(self, "שגיאה", "נתיב קובץ לא תקין")
            return False
        
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "שגיאה", "יש לבחור קובץ טקסט (txt) בלבד")
            return False
        
        self.file_path = file_path
        return True

    def run_script(self):
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return

        replace_type = self.replace_type.currentText()

        try:
            replacements_made = self.update_file(self.file_path, replace_type)
            
            if replacements_made > 0:
                QMessageBox.information(
                    self, 
                    "!מזל טוב", 
                    f"בוצעו {replacements_made} החלפות בהצלחה!"
                )
                self.changes_made.emit()  # שליחת סיגנל על שינויים
            else:
                QMessageBox.information(
                    self, 
                    "!שים לב", 
                    "לא נמצאו כותרות להחלפה"
                )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "שגיאה", 
                f"אירעה שגיאה בעיבוד הקובץ: {str(e)}"
            )

    def update_file(self, file_path, replace_type):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            previous_title = ""
            previous_level = ""
            replacements_made = 0  # ספירת כמות ההחלפות

            def replace_match(match):
                nonlocal previous_title, previous_level, replacements_made
                level = match.group(1)
                title = match.group(2)

                # בדיקה אם הכותרת היא "דף"
                if re.match(r"דף \S+\.?", title):
                    previous_title = title.strip()
                    previous_level = level
                    return match.group(0)

                # בדיקה אם הכותרת היא "עמוד ב"
                elif title == "עמוד ב":
                    replacements_made += 1  # הוחלפה כותרת
                    if replace_type == "נקודותיים":
                        return f'<h{previous_level}>{previous_title.rstrip(".")}:</h{previous_level}>'
                    elif replace_type == "ע\"ב":
                        # הסרת "ע\"א" או "עמוד א" מהכותרת הקודמת אם קיימים
                        modified_title = re.sub(r'( ע\"א| עמוד א)$', '', previous_title)
                        return f'<h{previous_level}>{modified_title.rstrip(".")} ע\"ב</h{previous_level}>'

                # אם זה לא אחד המקרים למעלה, נשאיר את הכותרת כפי שהיא
                return match.group(0)

            content = re.sub(r'<h([1-9])>(.*?)</h\1>', replace_match, content)

            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(content)

            return replacements_made

        except FileNotFoundError:
            QMessageBox.critical(self, "שגיאה", "הקובץ לא נמצא")
            return 0
        except UnicodeDecodeError:
            QMessageBox.critical(self, "שגיאה", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return 0
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")
            return 0

    def load_icon_from_base64(self, base64_string):
        """טעינת אייקון ממחרוזת Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)







    
   
def create_labeled_widget(label_text, widget):
    """יוצר widget עם תווית"""
    container = QWidget()
    v_layout = QVBoxLayout()
    v_layout.setContentsMargins(0, 0, 0, 0)
    v_layout.setSpacing(2)
    label = QLabel(label_text)
    label.setStyleSheet("font-size: 18px;")
    v_layout.addWidget(label)
    v_layout.addWidget(widget)
    container.setLayout(v_layout)
    return container

class בדיקת_שגיאות_בכותרות(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.setWindowTitle("בדיקת שגיאות בכותרות")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        # תווים בתחילת וסוף הכותרת
        regex_layout = QHBoxLayout()
        
        re_start_label = QLabel("תו/ים בתחילת הכותרת:")
        re_start_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        re_end_label = QLabel("תו/ים בסוף הכותרת:")
        self.re_end_entry = QLineEdit()
        self.re_end_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        self.gershayim_var = QCheckBox("כולל גרשיים")
        
        regex_layout.addWidget(self.gershayim_var)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_start_label)
        layout.addLayout(regex_layout)

        # יצירת תיבות טקסט להצגת תוצאות
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setReadOnly(True)
        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setReadOnly(True)

        # יצירת מכולות עם תוויות
        regex_container = create_labeled_widget(
            "פירוט הכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למעלה)\n"
            "אם יש רווח לפני או אחרי הכותרת, זה גם יוצג כשגיאה",
            self.unmatched_regex_text
        )
        tags_container = create_labeled_widget(
            "פירוט הכותרות שאינן לפי הסדר",
            self.unmatched_tags_text
        )

        # יצירת מפריד אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)
        v_splitter.addWidget(regex_container)
        v_splitter.addWidget(tags_container)
        layout.addWidget(v_splitter)

        self.setLayout(layout)

    def load_file_and_process(self, file_path):
        """עיבוד הקובץ והצגת התוצאות"""
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                html_content = file.read()
            
            re_start = self.re_start_entry.text()
            re_end = self.re_end_entry.text()
            gershayim = self.gershayim_var.isChecked()

            unmatched_regex, unmatched_tags = self.process_html(html_content, re_start, re_end, gershayim)
            
            # הצגת התוצאות
            if unmatched_regex:
                self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
            else:
                self.unmatched_regex_text.setPlainText("לא נמצאו שגיאות")
                
            if unmatched_tags:
                self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))
            else:
                self.unmatched_tags_text.setPlainText("לא נמצאו שגיאות")

        except Exception as e:
            QMessageBox.critical(None, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def process_html(self, html_content, re_start, re_end, gershayim):
        """עיבוד תוכן ה-HTML ובדיקת שגיאות"""
        soup = BeautifulSoup(html_content, 'html.parser')

        # יצירת תבנית regex
        if re_start and re_end:
            pattern = re.compile(f"^{re_start}.+[{re_end}]$")
        elif re_start:
            pattern = re.compile(f"^{re_start}.+['א-ת]$")
        elif re_end:
            pattern = re.compile(f"^[א-ת].+[{re_end}]$")
        else:
            pattern = re.compile(r"^[א-ת].+[א-ת']$")

        unmatched_regex = []
        unmatched_tags = []

        # בדיקת כותרות h2-h6
        for level in range(2, 7):
            headers = soup.find_all(f"h{level}")
            
            if not headers:
                unmatched_tags.append(f"מידע: אין בקובץ כותרות ברמה {level}")
                continue

            for i in range(len(headers) - 1):
                curr_header = headers[i].string or ""
                next_header = headers[i + 1].string or ""
                
                if not curr_header or not next_header:
                    continue

                # בדיקת תבנית
                if not re.match(pattern, curr_header):
                    unmatched_regex.append(curr_header)

                # חילוץ המספר מהכותרת
                curr_parts = curr_header.split()
                next_parts = next_header.split()
                
                curr_num = curr_parts[1] if len(curr_parts) > 1 else curr_header
                next_num = next_parts[1] if len(next_parts) > 1 else next_header

                # בדיקת גרשיים
                if gershayim:
                    if gematria.to_number(curr_num) <= 9:
                        if "'" not in curr_num:
                            unmatched_tags.append(curr_num)
                    elif '"' not in curr_num:
                        unmatched_tags.append(curr_num)
                elif "'" in curr_num or '"' in curr_num:
                    unmatched_tags.append(curr_num)

                # בדיקת רצף
                if not gematria.to_number(curr_num) + 1 == gematria.to_number(next_num):
                    unmatched_tags.append(f"כותרת נוכחית - {curr_header}, כותרת הבאה - {next_header}")

            # בדיקת הכותרת האחרונה
            if headers:
                last_header = headers[-1].string or ""
                if last_header and not re.match(pattern, last_header):
                    unmatched_regex.append(last_header)

                last_num = last_header.split()[1] if len(last_header.split()) > 1 else last_header
                if gershayim:
                    if gematria.to_number(last_num) <= 9:
                        if "'" not in last_num:
                            unmatched_tags.append(last_num)
                    elif '"' not in last_num:
                        unmatched_tags.append(last_num)
                elif "'" in last_num or '"' in last_num:
                    unmatched_tags.append(last_num)

        return unmatched_regex, unmatched_tags
    
class בדיקת_שגיאות_בתגים(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.setWindowTitle("בודק שגיאות בעיצוב")
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # יצירת תיבות טקסט
        self.opening_without_closing = QTextEdit()
        self.opening_without_closing.setReadOnly(True)

        self.closing_without_opening = QTextEdit()
        self.closing_without_opening.setReadOnly(True)

        self.heading_errors = QTextEdit()
        self.heading_errors.setReadOnly(True)

        # יצירת מכולות עם תוויות
        opening_container = create_labeled_widget(
            "תגים פותחים ללא תגים סוגרים",
            self.opening_without_closing
        )
        closing_container = create_labeled_widget(
            "תגים סוגרים ללא תגים פותחים",
            self.closing_without_opening
        )
        heading_container = create_labeled_widget(
            "טקסט שאינו חלק מכותרת, שנמצא באותה שורה עם הכותרת",
            self.heading_errors
        )

        # יצירת מפריד אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)
        v_splitter.addWidget(opening_container)
        v_splitter.addWidget(closing_container)
        v_splitter.addWidget(heading_container)

        main_layout.addWidget(v_splitter)
        self.setLayout(main_layout)

    def load_file_and_check(self, file_path):
        """בדיקת שגיאות בקובץ"""
        # ניקוי תוצאות קודמות
        self.opening_without_closing.clear()
        self.closing_without_opening.clear()
        self.heading_errors.clear()

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            opening_without_closing_list = []
            closing_without_opening_list = []
            heading_errors_list = []

            for line_num, line in enumerate(lines, 1):
                # בדיקת תגים
                tags_in_line = re.findall(r'<(/?\w+)>', line)
                stack = []

                for tag in tags_in_line:
                    if not tag.startswith('/'):  # תג פותח
                        stack.append(tag)
                    else:  # תג סוגר
                        if stack and stack[-1] == tag[1:]:
                            stack.pop()
                        else:
                            closing_without_opening_list.append(
                                f"שורה {line_num}: </{tag[1:]}> || {line.strip()}"
                            )

                # תגים שנשארו פתוחים
                for tag in stack:
                    opening_without_closing_list.append(
                        f"שורה {line_num}: <{tag}> || {line.strip()}"
                    )

                # בדיקת טקסט מחוץ לכותרות
                for tag in ["h2", "h3", "h4", "h5", "h6"]:
                    heading_pattern = rf'<{tag}>.*?</{tag}>'
                    match = re.search(heading_pattern, line)
                    if match:
                        start, end = match.span()
                        before = line[:start].strip()
                        after = line[end:].strip()
                        if before or after:
                            heading_errors_list.append(f"שורה {line_num}: {line.strip()}")

            # הצגת תוצאות
            self.opening_without_closing.setPlainText(
                "\n".join(opening_without_closing_list) if opening_without_closing_list 
                else "לא נמצאו שגיאות"
            )
            
            self.closing_without_opening.setPlainText(
                "\n".join(closing_without_opening_list) if closing_without_opening_list 
                else "לא נמצאו שגיאות"
            )
            
            self.heading_errors.setPlainText(
                "\n".join(heading_errors_list) if heading_errors_list 
                else "לא נמצאו שגיאות"
            )

        except Exception as e:
            QMessageBox.critical(None, "שגיאה", f"שגיאה בבדיקת הקובץ: {str(e)}")


class CheckHeadingErrorsOriginal(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("בודק כותרות + בודק תגים ביחד")
        self.setWindowIcon(self.get_app_icon())
        self.resize(1250, 700)

        # יצירת הווידג'טים המשניים
        self.check_headings_widget = בדיקת_שגיאות_בכותרות()
        self.html_tag_checker_widget = בדיקת_שגיאות_בתגים()
        
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # יצירת מפריד אופקי
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle:horizontal {
                width: 5px;
                margin: 1.5px;
                background: gray;
            }
        """)
        
        splitter.setChildrenCollapsible(False)
        
        # הגדרת מינימום רוחב
        self.html_tag_checker_widget.setMinimumWidth(10)
        self.check_headings_widget.setMinimumWidth(10)

        # הוספת הווידג'טים למיכל
        html_container = QWidget()
        self.html_container_layout = QVBoxLayout(html_container)
        self.html_container_layout.setContentsMargins(0, 0, 0, 0)
        self.html_container_layout.addWidget(self.html_tag_checker_widget)

        # תווית לציורים בספר
        self.pic_count_label = QLabel("")
        self.pic_count_label.setStyleSheet("font-size: 18px; color: blue;")
        self.pic_count_label.setVisible(False)
        
        splitter.addWidget(html_container)
        splitter.addWidget(self.check_headings_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        layout.addWidget(splitter)
        self.setLayout(layout)

    def set_file_path(self, file_path):
        """קבלת נתיב הקובץ ועיבודו"""
        self.file_path = file_path
        self.process_file(file_path)

    def process_file(self, file_path):
        """עיבוד הקובץ ובדיקת שגיאות"""
        try:
            # הפעלת בדיקות בשני הווידג'טים
            self.check_headings_widget.load_file_and_process(file_path)
            self.html_tag_checker_widget.load_file_and_check(file_path)

            # בדיקת ציורים בספר
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                count = content.count("ציור בספר")
                
                if count > 0:
                    text = (f'שים לב! יש בספר {count} ציורים.\n'
                           'חפש בתוך הספר את המילים "ציור בספר",\n'
                           'הורד את הספר מהיברובוקס, עשה צילום מסך לתמונה,\n'
                           'והמר אותה לטקסט ע"י תוכנה מספר 10')
                    self.pic_count_label.setText(text)
                    self.pic_count_label.setVisible(True)
                    if self.pic_count_label.parent() is None:
                        self.html_container_layout.addWidget(self.pic_count_label)
                else:
                    self.pic_count_label.setVisible(False)

        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def get_app_icon(self):
        """טעינת אייקון מקידוד Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(icon_base64))
        return QIcon(pixmap)
    
# ==========================================
# Script 8: בדיקת שגיאות בכותרות מותאם לספרים על השס (9 לשעבר)
# ==========================================

def create_labeled_widget(label_text, widget):
    """יוצר widget עם תווית בעיצוב אחיד"""
    container = QWidget()
    v_layout = QVBoxLayout()
    v_layout.setContentsMargins(0, 0, 0, 0)
    v_layout.setSpacing(5)
    
    label = QLabel(label_text)
    label.setStyleSheet("""
        QLabel {
            color: #1a365d;
            font-family: "Segoe UI", Arial;
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 5px;
        }
    """)
    
    # עיצוב ל-QTextEdit
    widget.setStyleSheet("""
        QTextEdit {
            border: 2px solid #2b4c7e;
            border-radius: 15px;
            padding: 10px;
            background-color: white;
            font-family: "Segoe UI", Arial;
            font-size: 12px;
        }
    """)
    
    v_layout.addWidget(label)
    v_layout.addWidget(widget)
    container.setLayout(v_layout)
    return container

class בדיקת_שגיאות_בכותרות_לשס(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.setWindowTitle("בדיקת שגיאות בכותרות לש\"ס")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        # תווים בתחילת וסוף הכותרת
        regex_layout = QHBoxLayout()
        
        re_start_label = QLabel("תו/ים בתחילת הכותרת:")
        re_start_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        re_end_label = QLabel("תו/ים בסוף הכותרת:")
        self.re_end_entry = QLineEdit()
        self.re_end_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_end_entry.setText('. :')
        
        self.gershayim_var = QCheckBox("כולל גרשיים")
        
        regex_layout.addWidget(self.gershayim_var)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_start_label)
        layout.addLayout(regex_layout)

        # יצירת תיבות טקסט להצגת תוצאות
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setReadOnly(True)
        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setReadOnly(True)

        # יצירת מכולות עם תוויות
        regex_container = create_labeled_widget(
            "פירוט הכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למעלה)\n"
            "אם יש רווח לפני או אחרי הכותרת, זה גם יוצג כשגיאה",
            self.unmatched_regex_text
        )
        tags_container = create_labeled_widget(
            "פירוט הכותרות שאינן לפי הסדר\n"
            "התוכנה מדלגת בבדיקה בכל פעם על כותרת אחת, בגלל הכותרות הכפולות לעמוד ב",
            self.unmatched_tags_text
        )

        # יצירת מפריד אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)
        v_splitter.addWidget(regex_container)
        v_splitter.addWidget(tags_container)
        layout.addWidget(v_splitter)

        self.setLayout(layout)

    def load_file_and_process(self, file_path):
        """עיבוד הקובץ והצגת התוצאות"""
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                html_content = file.read()
            
            re_start = self.re_start_entry.text()
            re_end = self.re_end_entry.text()
            gershayim = self.gershayim_var.isChecked()

            unmatched_regex, unmatched_tags = self.process_html(html_content, re_start, re_end, gershayim)
            
            # הצגת התוצאות
            if unmatched_regex:
                self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
            else:
                self.unmatched_regex_text.setPlainText("לא נמצאו שגיאות")
                
            if unmatched_tags:
                self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))
            else:
                self.unmatched_tags_text.setPlainText("לא נמצאו שגיאות")

        except Exception as e:
            QMessageBox.critical(None, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def process_html(self, html_content, re_start, re_end, gershayim):
        """עיבוד תוכן ה-HTML ובדיקת שגיאות"""
        soup = BeautifulSoup(html_content, 'html.parser')

        # יצירת תבנית regex
        if re_start and re_end:
            pattern = re.compile(f"^{re_start}.+[{re_end}]$")
        elif re_start:
            pattern = re.compile(f"^{re_start}.+['א-ת]$")
        elif re_end:
            pattern = re.compile(f"^[א-ת].+[{re_end}]$")
        else:
            pattern = re.compile(r"^[א-ת].+[א-ת']$")

        unmatched_regex = []
        unmatched_tags = []

        # בדיקת כותרות h2-h6
        for level in range(2, 7):
            headers = soup.find_all(f"h{level}")
            
            if not headers:
                unmatched_tags.append(f"מידע: אין בקובץ כותרות ברמה {level}")
                continue

            # עיבוד כל הכותרות למעט שתי האחרונות
            for i in range(len(headers) - 2):
                curr_header = headers[i].string or ""
                next_header = headers[i + 2].string or ""  # דילוג על כותרת אחת
                
                if not curr_header or not next_header:
                    continue

                # בדיקת תבנית
                if not re.match(pattern, curr_header):
                    unmatched_regex.append(curr_header)

                # חילוץ המספר מהכותרת
                curr_parts = curr_header.split()
                next_parts = next_header.split()
                
                curr_num = curr_parts[1] if len(curr_parts) > 1 else curr_header
                next_num = next_parts[1] if len(next_parts) > 1 else next_header

                # בדיקת גרשיים
                if gershayim:
                    if gematria.to_number(curr_num) <= 9:
                        if "'" not in curr_num:
                            unmatched_tags.append(curr_num)
                    elif '"' not in curr_num:
                        unmatched_tags.append(curr_num)
                elif "'" in curr_num or '"' in curr_num:
                    unmatched_tags.append(curr_num)

                # בדיקת רצף (עם דילוג על כותרת אחת)
                if not gematria.to_number(curr_num) + 1 == gematria.to_number(next_num):
                    unmatched_tags.append(f"כותרת נוכחית - {curr_header}, כותרת הבאה - {next_header}")

            # בדיקת שתי הכותרות האחרונות
            if len(headers) >= 2:
                for last_header in [headers[-2].string or "", headers[-1].string or ""]:
                    if last_header and not re.match(pattern, last_header):
                        unmatched_regex.append(last_header)
                    
                    parts = last_header.split()
                    last_num = parts[1] if len(parts) > 1 else last_header
                    
                    if gershayim:
                        if gematria.to_number(last_num) <= 9:
                            if "'" not in last_num:
                                unmatched_tags.append(last_num)
                        elif '"' not in last_num:
                            unmatched_tags.append(last_num)
                    elif "'" in last_num or '"' in last_num:
                        unmatched_tags.append(last_num)

        return unmatched_regex, unmatched_tags
class בדיקת_שגיאות_בתגים_לשס(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.setWindowTitle("בודק שגיאות בעיצוב לש\"ס")
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # יצירת תיבות טקסט
        self.opening_without_closing = QTextEdit()
        self.opening_without_closing.setReadOnly(True)

        self.closing_without_opening = QTextEdit()
        self.closing_without_opening.setReadOnly(True)

        self.heading_errors = QTextEdit()
        self.heading_errors.setReadOnly(True)

        # יצירת מכולות עם תוויות
        opening_container = create_labeled_widget(
            "תגים פותחים ללא תגים סוגרים",
            self.opening_without_closing
        )
        closing_container = create_labeled_widget(
            "תגים סוגרים ללא תגים פותחים",
            self.closing_without_opening
        )
        heading_container = create_labeled_widget(
            "טקסט שאינו חלק מכותרת, שנמצא באותה שורה עם הכותרת",
            self.heading_errors
        )

        # יצירת מפריד אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)
        v_splitter.addWidget(opening_container)
        v_splitter.addWidget(closing_container)
        v_splitter.addWidget(heading_container)

        main_layout.addWidget(v_splitter)
        self.setLayout(main_layout)

    def load_file_and_check(self, file_path):
        """בדיקת שגיאות בקובץ"""
        # ניקוי תוצאות קודמות
        self.opening_without_closing.clear()
        self.closing_without_opening.clear()
        self.heading_errors.clear()

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            opening_without_closing_list = []
            closing_without_opening_list = []
            heading_errors_list = []

            for line_num, line in enumerate(lines, 1):
                # בדיקת תגים
                tags_in_line = re.findall(r'<(/?\w+)>', line)
                stack = []

                for tag in tags_in_line:
                    if not tag.startswith('/'):  # תג פותח
                        stack.append(tag)
                    else:  # תג סוגר
                        if stack and stack[-1] == tag[1:]:
                            stack.pop()
                        else:
                            closing_without_opening_list.append(
                                f"שורה {line_num}: </{tag[1:]}> || {line.strip()}"
                            )

                # תגים שנשארו פתוחים
                for tag in stack:
                    opening_without_closing_list.append(
                        f"שורה {line_num}: <{tag}> || {line.strip()}"
                    )

                # בדיקת טקסט מחוץ לכותרות
                for tag in ["h2", "h3", "h4", "h5", "h6"]:
                    heading_pattern = rf'<{tag}>.*?</{tag}>'
                    match = re.search(heading_pattern, line)
                    if match:
                        start, end = match.span()
                        before = line[:start].strip()
                        after = line[end:].strip()
                        if before or after:
                            heading_errors_list.append(f"שורה {line_num}: {line.strip()}")

            # הצגת תוצאות
            self.opening_without_closing.setPlainText(
                "\n".join(opening_without_closing_list) if opening_without_closing_list 
                else "לא נמצאו שגיאות"
            )
            
            self.closing_without_opening.setPlainText(
                "\n".join(closing_without_opening_list) if closing_without_opening_list 
                else "לא נמצאו שגיאות"
            )
            
            self.heading_errors.setPlainText(
                "\n".join(heading_errors_list) if heading_errors_list 
                else "לא נמצאו שגיאות"
            )

        except Exception as e:
            QMessageBox.critical(None, "שגיאה", f"שגיאה בבדיקת הקובץ: {str(e)}")


class CheckHeadingErrorsCustom(QWidget):
    changes_made = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.setWindowTitle("בודק כותרות + בודק תגים לש\"ס")
        self.setWindowIcon(self.get_app_icon())
        self.resize(1250, 700)

        # יצירת הווידג'טים המשניים
        self.check_headings_widget = בדיקת_שגיאות_בכותרות_לשס()
        self.html_tag_checker_widget = בדיקת_שגיאות_בתגים_לשס()
        
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # יצירת מפריד אופקי
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle:horizontal {
                width: 5px;
                margin: 1.5px;
                background: gray;
            }
        """)
        
        splitter.setChildrenCollapsible(False)
        
        # הגדרת מינימום רוחב
        self.html_tag_checker_widget.setMinimumWidth(10)
        self.check_headings_widget.setMinimumWidth(10)

        # הוספת הווידג'טים למיכל
        html_container = QWidget()
        self.html_container_layout = QVBoxLayout(html_container)
        self.html_container_layout.setContentsMargins(0, 0, 0, 0)
        self.html_container_layout.addWidget(self.html_tag_checker_widget)

        # תווית לציורים בספר
        self.pic_count_label = QLabel("")
        self.pic_count_label.setStyleSheet("font-size: 18px; color: blue;")
        self.pic_count_label.setVisible(False)
        
        splitter.addWidget(html_container)
        splitter.addWidget(self.check_headings_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        layout.addWidget(splitter)
        self.setLayout(layout)

    def set_file_path(self, file_path):
        """קבלת נתיב הקובץ ועיבודו"""
        self.file_path = file_path
        self.process_file(file_path)

    def process_file(self, file_path):
        """עיבוד הקובץ ובדיקת שגיאות"""
        try:
            # הפעלת בדיקות בשני הווידג'טים
            self.check_headings_widget.load_file_and_process(file_path)
            self.html_tag_checker_widget.load_file_and_check(file_path)

            # בדיקת ציורים בספר
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                count = content.count("ציור בספר")
                
                if count > 0:
                    text = (f'שים לב! יש בספר {count} ציורים.\n'
                           'חפש בתוך הספר את המילים "ציור בספר",\n'
                           'הורד את הספר מהיברובוקס, עשה צילום מסך לתמונה,\n'
                           'והמר אותה לטקסט ע"י תוכנה מספר 10')
                    self.pic_count_label.setText(text)
                    self.pic_count_label.setVisible(True)
                    if self.pic_count_label.parent() is None:
                        self.html_container_layout.addWidget(self.pic_count_label)
                else:
                    self.pic_count_label.setVisible(False)

        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def get_app_icon(self):
        """טעינת אייקון מקידוד Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(icon_base64))
        return QIcon(pixmap)    
# ==========================================
# Script 9: המרת תמונה לטקסט 10(לשעבר)
# ==========================================

class ImageToHtmlApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = "" 
        self.setWindowTitle("המרת תמונה לטקסט")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        #self.setGeometry(100, 100, 350,350)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)  
        
        # עיצוב תוויות
        label_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 10px;
            }
        """

        # עיצוב כפתורים
        button_style = """
           QPushButton {
               border-radius: 15px;
               padding: 5px;
               background-color: #eaeaea;
               color: black;
               font-weight: bold;
               font-family: "Segoe UI", Arial;
               font-size: 11px;
               min-height: 30px;
              max-height: 30px;
             }
           QPushButton:hover {
               background-color: #b7b5b5;
           }
           QPushButton:disabled {
               background-color: #cccccc;
               color: #666666;
             }
          """

        # תווית מידע
        self.information_label = QLabel("לפניך מספר אפשרויות לבחירת התמונה\nבחר אחת מהן")
        self.information_label.setAlignment(Qt.AlignCenter)
        self.information_label.setStyleSheet(label_style)
        self.layout.addWidget(self.information_label)

        # אזור גרירה
        self.label = QLabel("גרור ושחרר את הקובץ", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            QLabel {
                border: 2px dashed #2b4c7e;
                border-radius: 15px;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 40px;
                background-color: #f8f9fa;
            }
        """)
        self.layout.addWidget(self.label)

        # תווית הוראות
        self.instruction_label = QtWidgets.QLabel("הדבק נתיב קובץ [או קישור מקוון לתמונה]\nאו הדבק את התמונה (Ctrl+V):")
        self.instruction_label.setAlignment(Qt.AlignCenter)
        self.instruction_label.setStyleSheet(label_style)
        self.layout.addWidget(self.instruction_label)

        # תיבת טקסט
        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.setStyleSheet("""
            QLineEdit {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 10px;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                min-height: 20px;
            }
            QLineEdit:focus {
                border-color: #1a73e8;
            }
        """)
        self.url_edit.textChanged.connect(self.on_text_changed)
        self.url_edit.returnPressed.connect(self.convert_image)
        self.layout.addWidget(self.url_edit)

        # כפתורים
        # כפתור עיון
        browse_container = QHBoxLayout()
        self.add_files_button = QPushButton('עיון', self)
        self.add_files_button.setStyleSheet(button_style)
        self.add_files_button.setFixedSize(80, 30)
        self.add_files_button.clicked.connect(self.open_file_dialog)
        browse_container.addStretch(1)
        browse_container.addWidget(self.add_files_button)
        browse_container.addStretch(1)
        self.layout.addLayout(browse_container)

        # כפתור המרה
        convert_container = QHBoxLayout()
        self.convert_btn = QtWidgets.QPushButton("המר")
        self.convert_btn.setEnabled(False)
        self.convert_btn.setStyleSheet(button_style)
        self.convert_btn.setFixedSize(80, 30)
        self.convert_btn.clicked.connect(self.convert_image)
        convert_container.addStretch(1)
        convert_container.addWidget(self.convert_btn)
        convert_container.addStretch(1)
        self.layout.addLayout(convert_container)

        # תוויות תוצאה
        success_label_style = """
            QLabel {
                color: #28a745;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 10px;
                background-color: #e8f5e9;
                border-radius: 10px;
            }
        """

        self.nextInFocusChain = QLabel("ההמרה בוצעה בהצלחה!")
        self.nextInFocusChain.setVisible(False)
        self.nextInFocusChain.setAlignment(Qt.AlignCenter)
        self.nextInFocusChain.setStyleSheet(success_label_style)
        self.layout.addWidget(self.nextInFocusChain)
        
        # כפתור העתקה
        copy_container = QHBoxLayout()
        self.copy_btn = QtWidgets.QPushButton("העתק")
        self.copy_btn.setEnabled(False)
        self.copy_btn.setStyleSheet(button_style)
        self.copy_btn.setFixedSize(80, 30)
        self.copy_btn.clicked.connect(self.copy_to_clipboard)
        copy_container.addStretch(1)
        copy_container.addWidget(self.copy_btn)
        copy_container.addStretch(1)
        self.layout.addLayout(copy_container)

        # תווית אישור העתקה
        self.cop = QLabel("הטקסט הועתק ללוח, ניתן להדביקו במסמך")  # הוספת התווית החסרה
        self.cop.setVisible(False)
        self.cop.setAlignment(Qt.AlignCenter)
        self.cop.setStyleSheet(success_label_style)
        self.layout.addWidget(self.cop)        

        # כפתור המרה חדשה
        new_convert_container = QHBoxLayout()
        self.convert_new_button = QPushButton('המרה חדשה', self)
        self.convert_new_button.setVisible(False)
        self.convert_new_button.setStyleSheet(button_style)
        self.convert_new_button.setFixedSize(100, 30)
        self.convert_new_button.clicked.connect(self.reset_for_new_convert)
        new_convert_container.addStretch(1)
        new_convert_container.addWidget(self.convert_new_button)
        new_convert_container.addStretch(1)
        self.layout.addLayout(new_convert_container)

        self.setAcceptDrops(True)
        self.img_data = None
        self.image_files = []

    def on_text_changed(self):
        text = self.url_edit.text().strip()
        if text.startswith("file:///"):
            text = text[8:]  # הסרת "file:///"
            self.url_edit.setText(text)  # עדכון השדה לאחר התיקון

        if os.path.exists(text):  # בדיקת קובץ מקומי
            self.label.setText("התמונה נטענה בהצלחה!")
            self.convert_btn.setEnabled(True)
        elif text.lower().startswith("http://") or text.lower().startswith("https://"):
            try:
                req = urllib.request.Request(text, method="HEAD")  # שליחה רק של בקשת HEAD לבדיקה
                urllib.request.urlopen(req)
                self.label.setText("הקישור תקין ונטען בהצלחה!")
                self.convert_btn.setEnabled(True)
            except Exception:
                self.label.setText("לא ניתן לפתוח את הקישור, ודא שהוא תקין")
                self.convert_btn.setEnabled(False)
        else:
            self.label.setText("הנתיב שסיפקת אינו קיים")
            self.convert_btn.setEnabled(False)

    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "בחר קבצי תמונה", "", 
                                                "קבצי תמונה (*.png;*.jpg;*.jpeg;*.svg;*.tif;*.heic;*.heif;*.ico;*.webp;*.tiff;*.gif;*.bmp)")
        if files:
            supported_extensions = ('.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff', '.heic', '.heif', '.ico', '.webp', '.gif', '.bmp')
            for file in files:
                if file.lower().endswith(supported_extensions):
                    self.image_files.append(file)
                    with open(file, 'rb') as f:
                        self.img_data = f.read()
                    self.label.setText("התמונה נטענה בהצלחה!")
                    self.convert_btn.setEnabled(True)
                else:
                    self.label.setText("הסיומת של הקובץ אינה נתמכת.\nבחר קובץ אחר")

    # פונקציה שמופעלת כשגוררים קובץ לחלון
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    # פונקציה שמופעלת כשמשחררים את הקבצים בחלון
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            supported_extensions = ('.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff', '.heic', '.heif', '.ico', '.webp', '.gif', '.bmp')
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path and os.path.exists(file_path):
                    if file_path.lower().endswith(supported_extensions):
                        with open(file_path, 'rb') as f:
                            self.img_data = f.read()
                        self.image_files.append(file_path)
                        self.label.setText("התמונה נטענה בהצלחה!")
                        self.convert_btn.setEnabled(True)
                    else:
                        self.label.setText("הסיומת של הקובץ אינה נתמכת.\nבחר קובץ אחר")

    def keyPressEvent(self, event):
        if event.matches(QtGui.QKeySequence.Paste):
            clipboard = QtWidgets.QApplication.clipboard()
            mime_data = clipboard.mimeData()
            # בדיקה אם מדובר בתמונה שהועתקה
            if mime_data.hasImage():
                image = clipboard.image()
                if not image.isNull():
                    buffer = QtCore.QBuffer()
                    buffer.open(QtCore.QBuffer.WriteOnly)
                    image.save(buffer, "PNG")
                    self.img_data = buffer.data().data()
                    self.label.setText("התמונה הודבקה בהצלחה!")
                    self.convert_btn.setEnabled(True)
            else:
                text = clipboard.text().strip().strip('"')
                self.url_edit.setText(text)
            event.accept()

    def convert_image(self):
        path = self.url_edit.text().strip().strip('"')
        if path.startswith("file:///"):  # טיפול בפרוטוקול file:///
            path = path[8:]  # הסרת "file:///"

        if not self.img_data and path:
            if path.lower().startswith("http://") or path.lower().startswith("https://"):
                try:
                    with urllib.request.urlopen(path) as resp:
                        self.img_data = resp.read()
                    self.label.setText("הקישור נטען בהצלחה!")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "שגיאה", f"לא ניתן לפתוח את הקישור:\n{e}")
                    return
            elif os.path.exists(path):  # בדיקה אם הקובץ קיים
                with open(path, 'rb') as f:
                    self.img_data = f.read()
                self.label.setText("התמונה נטענה בהצלחה!")
            else:
                QtWidgets.QMessageBox.warning(self, "שגיאה", "לא ניתן לאתר קובץ בנתיב שסיפקת, ודא שהנתיב נכון")
                return

        if not self.img_data:
            QtWidgets.QMessageBox.warning(self, "שגיאה", "לא נמצאה תמונה להמרה")
            return

        # זיהוי סוג הקובץ על בסיס הסיומת
        file_extension = "png"  # ברירת מחדל
        if self.image_files:
            file_extension = os.path.splitext(self.image_files[0])[-1].lstrip(".").lower()
        elif path:
            file_extension = os.path.splitext(path)[-1].lstrip(".").lower()

        encoded = base64.b64encode(self.img_data).decode('utf-8')
        self.html_code = f'<img src="data:image/{file_extension};base64,{encoded}" >'
        self.copy_btn.setEnabled(True)
        self.nextInFocusChain.setVisible(True)

    def copy_to_clipboard(self):
        clipboard = QtWidgets.QApplication.clipboard()
        clipboard.setText(self.html_code)
        self.cop.setVisible(True)
        self.show_post_convert_buttons()

    # פונקציה להצגת כפתורים אחרי ההמרה
    def show_post_convert_buttons(self):
        self.add_files_button.setEnabled(True)
        self.convert_btn.setEnabled(False)
        self.convert_new_button.setVisible(True)

    # פונקציה לאיפוס עבור המרת קבצים חדשים
    def reset_for_new_convert(self):
        self.img_data = None
        self.image_files = []
        self.label.setText("גרור ושחרר קבצי תמונה")
        self.convert_btn.setEnabled(False)
        self.convert_new_button.setVisible(False)
        self.nextInFocusChain.setVisible(False)
        self.copy_btn.setEnabled(False)
        self.cop.setVisible(False)
        self.url_edit.clear()

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# ==========================================
# Script 10: תיקון שגיאות נפוצות (11 לשעבר)
# ==========================================

class TextCleanerApp(QWidget):
    changes_made = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.file_path = ""
        self.originalText = ""
        self.setWindowTitle("תיקון שגיאות נפוצות")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        #self.setGeometry(100, 100, 500, 500)
        self.setFixedWidth(600)
        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # הסבר למשתמש בחלק העליון - באדום
        explanation = QLabel(
            "שים לב!\n\n"
            "התוכנה תיקון שגיאות נפוצות בטקסט.\n"
            "סמן את האפשרויות הרצויות ולחץ על 'הרץ כעת'.\n"
            "ניתן לבטל את השינוי האחרון באמצעות הכפתור 'בטל שינוי אחרון'."
        )
        explanation.setStyleSheet("""
            QLabel {
                color: #8B0000;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 20px;
                background-color: #FFE4E1;
                border: 2px solid #CD5C5C;
                border-radius: 15px;
                font-weight: bold;
            }
        """)
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # כפתורי בחירת הכל/ביטול הכל
        button_container = QHBoxLayout()
        
        self.selectAllBtn = QPushButton("בחר הכל")
        self.selectAllBtn.clicked.connect(self.selectAll)
        self.selectAllBtn.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
            QPushButton:pressed {
                background-color: #a0a0a0;
            }
        """)
        
        self.deselectAllBtn = QPushButton("בטל הכל")
        self.deselectAllBtn.clicked.connect(self.deselectAll)
        self.deselectAllBtn.setStyleSheet(self.selectAllBtn.styleSheet())
        
        button_container.addStretch(1)
        button_container.addWidget(self.selectAllBtn)
        button_container.addWidget(self.deselectAllBtn)
        button_container.addStretch(1)
        layout.addLayout(button_container)

        # קונטיינר לתיבות הסימון
        checkbox_style = """
            QCheckBox {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 5px;
                spacing: 10px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border: 2px solid #2b4c7e;
                border-radius: 5px;
            }
            QCheckBox::indicator:unchecked {
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #2b4c7e;
                image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3E%3Cpath fill='white' d='M3.5 8.5l3 3 6-6-1-1-5 5-2-2z'/%3E%3C/svg%3E");
            }
        """
        # תיבות סימון לאפשרויות שונות
        self.checkBoxes = {
            "remove_empty_lines": QCheckBox("מחיקת שורות ריקות"),
            "remove_double_spaces": QCheckBox("מחיקת רווחים כפולים"),
            "remove_spaces_before": QCheckBox("\u202Bמחיקת רווחים לפני - . , : ) ]"),
            "remove_spaces_after": QCheckBox("\u202Bמחיקת רווחים אחרי - [ ("),
            "remove_spaces_around_newlines": QCheckBox("מחיקת רווחים לפני ואחרי אנטר"),
            "replace_double_quotes": QCheckBox("החלפת 2 גרשים בודדים בגרשיים"),
            "normalize_quotes": QCheckBox("המרת גרשיים מוזרים לגרשיים רגילים"),
        }

        # הוספת תיבות הסימון לממשק
        checkbox_container = QVBoxLayout()
        for checkbox in self.checkBoxes.values():
            checkbox.setStyleSheet(checkbox_style)
            checkbox.setChecked(True)
            checkbox_container.addWidget(checkbox)
        layout.addLayout(checkbox_container)

        # כפתורי הפעלה וביטול
        action_buttons_container = QVBoxLayout()
        
        self.cleanBtn = QPushButton("הרץ כעת")
        self.cleanBtn.clicked.connect(self.runCleanText)
        self.cleanBtn.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #4CAF50;  /* ירוק */
                color: white;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        
        self.undoBtn = QPushButton("בטל שינוי אחרון")
        self.undoBtn.clicked.connect(self.undoChanges)
        self.undoBtn.setEnabled(False)
        self.undoBtn.setStyleSheet("""
            QPushButton {
                border-radius: 15px;
                padding: 5px;
                background-color: #f44336;  /* אדום */
                color: white;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                min-height: 30px;
                max-height: 30px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:pressed {
                background-color: #c41810;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        action_buttons_container.addWidget(self.cleanBtn, alignment=Qt.AlignCenter)
        action_buttons_container.addWidget(self.undoBtn, alignment=Qt.AlignCenter)
        layout.addLayout(action_buttons_container)

        # מרווח גמיש בתחתית
        layout.addStretch()

        self.setLayout(layout)


    def set_file_path(self, path):
        """מקבלת את נתיב הקובץ מהחלון הראשי"""
        self.file_path = path

    def runCleanText(self):
        """הפעלת פונקציית הניקוי"""
        if not self.file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.cleanText()

    def cleanText(self):
        """פונקציית הניקוי העיקרית"""
        try:
            # קריאת הקובץ
            with open(self.file_path, 'r', encoding='utf-8') as file:
                text = file.read()
            
            # שמירת הטקסט המקורי לצורך ביטול
            self.originalText = text
            self.undoBtn.setEnabled(True)
            
            # ביצוע כל הפעולות שנבחרו
            if self.checkBoxes["remove_empty_lines"].isChecked():
                text = re.sub(r'\n\s*\n', '\n', text)
            
            if self.checkBoxes["remove_double_spaces"].isChecked():
                text = re.sub(r' +', ' ', text)
            
            if self.checkBoxes["remove_spaces_before"].isChecked():
                text = re.sub(r'\s+([)\],.:])', r'\1', text)
            
            if self.checkBoxes["remove_spaces_after"].isChecked():
                text = re.sub(r'([\[(])\s+', r'\1', text)
            
            if self.checkBoxes["remove_spaces_around_newlines"].isChecked():
                text = re.sub(r'\s*\n\s*', '\n', text)
            
            if self.checkBoxes["replace_double_quotes"].isChecked():
                text = text.replace("''", '"').replace("``", '"').replace("''", '"')
            
            if self.checkBoxes["normalize_quotes"].isChecked():
                text = text.replace(""", '"').replace(""", '"').replace("'", "'").replace("'", "'").replace("„", '"')
            
            # מחיקת רווחים בסוף הקובץ
            text = text.rstrip()

            # בדיקה אם היו שינויים
            if text == self.originalText:
                QMessageBox.information(self, "שינויי טקסט", "אין מה להחליף בקובץ זה.")
                return

            # שמירת השינויים
            with open(self.file_path, 'w', encoding='utf-8') as file:
                file.write(text)
            
            QMessageBox.information(self, "שינויי טקסט", "השינויים בוצעו בהצלחה.")
            self.changes_made.emit()

        except FileNotFoundError:
            QMessageBox.critical(self, "שגיאה", "הקובץ לא נמצא")
        except UnicodeDecodeError:
            QMessageBox.critical(self, "שגיאה", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")
    
    def selectAll(self):
        """בחירת כל האפשרויות"""
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(True)
    
    def deselectAll(self):
        """ביטול כל האפשרויות"""
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(False)
    
    def undoChanges(self):
        """ביטול השינוי האחרון"""
        if not self.file_path or not self.originalText:
            return
            
        try:
            with open(self.file_path, 'w', encoding='utf-8') as file:
                file.write(self.originalText)
            QMessageBox.information(self, "ביטול שינויים", "השינויים בוטלו בהצלחה.")
            self.changes_made.emit()
            self.undoBtn.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בביטול השינויים: {str(e)}")

    def load_icon_from_base64(self, base64_string):
        """טעינת אייקון מקידוד Base64"""
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# ==========================================
# Main Menu: תפריט ראשי ופונקציות מרכזיות
# ==========================================
class MainMenu(QWidget):
    def __init__(self):
        super().__init__()
        self.document_history = []
        self.redo_history = []        
        self.current_file_path = ""
        self.current_index = -1        
        self.last_processor_title = ""
        self.current_version = "3.2"
        
        # הגדרת החלון
        self.setWindowTitle("עריכת ספרי דיקטה עבור אוצריא")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 1000, 600)
        self.init_ui()

        if sys.platform == 'win32':
            QtWin.setCurrentProcessExplicitAppUserModelID(myappid)

    def init_ui(self):
        main_layout = QHBoxLayout()
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # מיכל ימני לכפתורים
        right_container = QWidget()
        right_container.setFixedWidth(300)
        right_layout = QVBoxLayout(right_container)

        buttons_grid_widget = QWidget()
        buttons_grid_widget.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Expanding)
        grid_layout = QGridLayout(buttons_grid_widget)
        grid_layout.setSpacing(5)

        # פאנל טקסט
        text_container = QWidget()
        text_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        text_layout = QVBoxLayout(text_container)
        text_layout.setContentsMargins(15, 15, 0, 20)

        # יצירת כפתורי פעולה
        action_buttons_layout = QHBoxLayout()
        
        # כפתור ביטול
        self.undo_button = QPushButton("⟲")
        self.undo_button.setStyleSheet("""
            font-weight: bold; 
            font-size: 14pt;
            padding: 5px;
        """)
        self.undo_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.undo_button.clicked.connect(self.undo_action)
        self.undo_button.setFixedSize(40, 40)
        self.undo_button.setToolTip("בטל")
        self.undo_button.setEnabled(False)
        
        # כפתור חזרה
        self.redo_button = QPushButton("⟳")
        self.redo_button.setStyleSheet("""
            font-weight: bold; 
            font-size: 14pt;
            padding: 5px;
        """)
        self.redo_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.redo_button.clicked.connect(self.redo_action)
        self.redo_button.setFixedSize(40, 40)
        self.redo_button.setToolTip("חזור")
        self.redo_button.setEnabled(False)
        
        # כפתור שמירה
        self.save_button = QPushButton("🖫")
        self.save_button.setStyleSheet("""
            font-weight: bold; 
            font-size: 14pt;
            padding: 5px;
        """)
        self.save_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setFixedSize(40, 40)
        self.save_button.setToolTip("שמור")
        self.save_button.setEnabled(False)

        # תווית סטטוס
        self.status_label = QLabel("לא בוצעו עדיין פעולות")
        self.status_label.setStyleSheet("""
            color: #666666;
            font-size: 14px;
            padding: 5px 15px;
            background-color: transparent;
            border-radius: 10px;
        """)
        self.status_label.setAlignment(Qt.AlignCenter)

        # כפתור הוספת קובץ
        add_file_button = QPushButton("הוסף קובץ")
        add_file_button.setFixedSize(100, 40)
        add_file_button.setCursor(QCursor(Qt.PointingHandCursor))
        add_file_button.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                padding: 5px;
                background-color: #eaeaea;
                color: black;
                font-weight: bold;
                font-size: 8.5pt;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }
        """)
        add_file_button.clicked.connect(self.select_file)

        # הוספת כפתורי פעולה ללייאאוט
        action_buttons_layout.addWidget(add_file_button)
        action_buttons_layout.addStretch(1)
        action_buttons_layout.addWidget(self.status_label)
        action_buttons_layout.addStretch(1)
        action_buttons_layout.addWidget(self.undo_button)
        action_buttons_layout.addWidget(self.redo_button)
        action_buttons_layout.addWidget(self.save_button)

        # יצירת תיבות הטקסט עם הטקסט הדוגמה
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setReadOnly(True)
        self.unmatched_regex_text.setHtml(DefaultTextContent.get_default_html())
        self.unmatched_regex_text.setStyleSheet(DefaultTextContent.get_default_style())

        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setReadOnly(True)
        self.unmatched_tags_text.setHtml(DefaultTextContent.get_default_html())
        self.unmatched_tags_text.setStyleSheet(DefaultTextContent.get_default_style())

        # תצוגת טקסט
        self.text_display = QtWidgets.QTextBrowser()
        self.text_display.setReadOnly(True)
        self.text_display.setLayoutDirection(Qt.RightToLeft)
        self.text_display.setStyleSheet("""
            QTextBrowser {
                background-color: transparent;
                border: 2px solid black;
                border-radius: 15px;
                padding: 20px 40px;
                font-family: "Frank Ruehl CLM",  "Segoe UI";
                font-size: 18px;
                line-height: 1.5;
            }
        """)

        # יצירת כפתורי עריכה
        text_bottom_buttons = QHBoxLayout()
        text_bottom_buttons.setSpacing(10)
        text_bottom_buttons.addSpacing(10)

        # הגדרת כפתורי עריכה
        buttons_data = [
            ("קטן", self.button1_function),
            ("גדול", self.button2_function),
            ("נטוי", self.button3_function),
            ("דגש", self.button4_function),
            ("H6", self.button5_function),
            ("H5", self.button6_function),
            ("H4", self.button7_function),
            ("H3", self.button8_function),
            ("H2", self.button9_function),
            ("H1", self.button10_function)
        ]

        button_style = """
            QPushButton {
                border-radius: 15px;
                padding: 6px 12px;
                background-color: #E8F0FE; 
                color: #1a365d;
                font-weight: bold;
                font-family: "Segoe UI", Arial;
                font-size: 7.5pt;
                min-width: 20px;
                min-height: 12px;
                width: 12px;
                height: 20px;
                border: 1px solid #c2d3f0;
            }
            QPushButton:hover {
                background-color: #d3e3fc; 
            }
            QPushButton:pressed {
                background-color: #bbd1f8;
                padding: 7px 11px 5px 13px;
            }
        """

        # יצירת כפתורי עריכה
        self.editing_buttons = []
        for button_text, func in reversed(buttons_data):
            button = QPushButton(button_text)
            button.setFixedSize(40, 40)
            button.setStyleSheet(button_style)
            button.setCursor(QCursor(Qt.PointingHandCursor))
            button.clicked.connect(func)
            button.hide()
            self.editing_buttons.insert(0, button)
            text_bottom_buttons.addWidget(button)

        text_bottom_buttons.addStretch(1)

        # כפתורי תפריט
        button_info = [
            ("1\n\nיצירת כותרות לאוצריא", self.open_create_headers_otzria),
            ("2\n\nיצירת כותרות לאותיות בודדות", self.open_create_single_letter_headers),
            ("3\n\nשינוי רמת כותרת", self.open_change_heading_level),
            ("4\n\nהדגשת מילה ראשונה וניקוד בסוף קטע", self.open_emphasize_and_punctuate),
            ("5\n\nיצירת כותרות לעמוד ב", self.open_create_page_b_headers),
            ("6\n\nהחלפת כותרות לעמוד ב", self.open_replace_page_b_headers),
            ("7\n\nבדיקת שגיאות", self.open_check_heading_errors_original),
            ("8\n בדיקת שגיאות לספרים על השס", self.open_check_heading_errors_custom),
            ("9\n\nהמרת תמונה לטקסט", self.open_Image_To_Html_App),
            ("10\n\nתיקון שגיאות נפוצות", self.open_Text_Cleaner_App),
        ]

        for i, (text, func) in enumerate(button_info):
            button = QPushButton(text)
            button.setFixedSize(250, 70)
            button.clicked.connect(func)
            button.setStyleSheet("""
                QPushButton {
                    border-radius: 30px;
                    padding: 10px;
                    margin: 5;
                    background-color: #eaeaea;
                    color: black;
                    font-weight: bold;
                    font-family: "Segoe UI", Arial;
                    font-size: 8.5pt;
                }
                QPushButton:hover {
                    background-color: #b7b5b5;
                }
            """)
            grid_layout.addWidget(button, i, 0)

        # כפתורים תחתונים
        bottom_buttons_layout = QHBoxLayout()
        
        about_button = QPushButton("i")
        about_button.setStyleSheet("font-weight: bold; font-size: 12pt;")
        about_button.setCursor(QCursor(Qt.PointingHandCursor))
        about_button.clicked.connect(self.open_about_dialog)
        about_button.setFixedSize(40, 40)
        
        update_button = QPushButton("⭳")
        update_button.setStyleSheet("""
            font-weight: bold; 
            font-size: 14pt;
        """)
        update_button.setCursor(QCursor(Qt.PointingHandCursor))
        update_button.clicked.connect(self.check_for_updates)
        update_button.setFixedSize(40, 40)
        update_button.setToolTip("עדכונים")

        self.edit_button = QPushButton("✍")
        self.edit_button.setStyleSheet("""
            font-weight: bold; 
            font-size: 14pt;
        """)
        self.edit_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.edit_button.setFixedSize(40, 40)
        self.edit_button.setToolTip("עריכה")
        self.edit_button.clicked.connect(self.edit_text)

        bottom_buttons_layout.addWidget(about_button)
        bottom_buttons_layout.addWidget(update_button)
        bottom_buttons_layout.addSpacing(120)
        bottom_buttons_layout.addWidget(self.edit_button)
        bottom_buttons_layout.addStretch()

        # סידור סופי של הלייאאוטים
        text_layout.insertLayout(0, action_buttons_layout)
        text_layout.addLayout(text_bottom_buttons)
        text_layout.addWidget(self.text_display)
       

        right_layout.addWidget(buttons_grid_widget)
        right_layout.addLayout(bottom_buttons_layout)

        main_layout.addWidget(right_container)
        main_layout.addWidget(text_container)
        main_layout.setStretchFactor(right_container, 1)
        main_layout.setStretchFactor(text_container, 2)

        self.setLayout(main_layout)
        
       

    def button1_function(self):
        """הקטנת הטקסט המסומן"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<span style="font-size: smaller;">{selected_text}</span>')
            self._safe_update_history(self.text_display.toHtml(), "קטן")

    def button2_function(self):
        """הגדלת הטקסט המסומן"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<span style="font-size: larger;">{selected_text}</span>')
            self._safe_update_history(self.text_display.toHtml(), "גדול")

    def button3_function(self):
        """הפיכת הטקסט לנטוי"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<i>{selected_text}</i>')
            self._safe_update_history(self.text_display.toHtml(), "נטוי")

    def button4_function(self):
        """הדגשת הטקסט"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<b>{selected_text}</b>')
            self._safe_update_history(self.text_display.toHtml(), "דגש")

    def button5_function(self):
        """הוספת כותרת H6"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h6>{selected_text}</h6>')
            self._safe_update_history(self.text_display.toHtml(), "הH6")

    def button6_function(self):
        """הוספת כותרת H5"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h5>{selected_text}</h5>')
            self._safe_update_history(self.text_display.toHtml(), "H5")

    def button7_function(self):
        """הוספת כותרת H4"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h4>{selected_text}</h4>')
            self._safe_update_history(self.text_display.toHtml(), "הH4")

    def button8_function(self):
        """הוספת כותרת H3"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h3>{selected_text}</h3>')
            self._safe_update_history(self.text_display.toHtml(), "H3")

    def button9_function(self):
        """הוספת כותרת H2"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h2>{selected_text}</h2>')
            self._safe_update_history(self.text_display.toHtml(), "H2")

    def button10_function(self):
        """הוספת כותרת H1"""
        cursor = self.text_display.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.removeSelectedText()
            cursor.insertHtml(f'<h1>{selected_text}</h1>')
            self._safe_update_history(self.text_display.toHtml(), "H1")   
        
    def edit_text(self):
        """פונקציה לניהול מצב העריכה בפאנל הקיים"""
        is_editing_mode = self.editing_buttons[0].isHidden()
        
        # שינוי מצב העריכה של פאנל הטקסט
        self.text_display.setReadOnly(not is_editing_mode)
        
        # שמירת המצב הנוכחי להיסטוריה אם נכנסים למצב עריכה
        if is_editing_mode:
            current_text = self.text_display.toPlainText()
            self._safe_update_history(current_text, "כניסה למצב עריכה")
            
            # סגנון למצב עריכה
            self.text_display.setStyleSheet("""
                QTextEdit {
                    background-color: white;
                    border: 2px solid #2b4c7e;
                    border-radius: 15px;
                    padding: 20px 40px;
                    font-family: "Segoe UI", Arial;
                    font-size: 14px;
                    line-height: 1.5;
                }
            """)
            
            # עדכון ממשק המשתמש למצב עריכה
            self.edit_button.setText("✗")
            self.edit_button.setToolTip("סגור מצב עריכה")
            self.status_label.setText("מצב עריכה פעיל")
            self.status_label.setStyleSheet("""
                color: #2b4c7e;
                font-size: 14px;
                padding: 5px 15px;
                background-color: #E8F0FE;
                border-radius: 10px;
            """)
            
            # הצגת כפתורי העריכה
            for button in self.editing_buttons:
                button.show()
        
        else:
            # שמירת השינויים לפני יציאה ממצב עריכה
            if self.current_file_path:
                current_text = self.text_display.toPlainText()
                self._safe_update_history(current_text, "יציאה ממצב עריכה")
            
            # סגנון למצב תצוגה
            self.text_display.setStyleSheet("""
                QTextEdit {
                    background-color: transparent;
                    border: 2px solid black;
                    border-radius: 15px;
                    padding: 20px 40px;
                    font-family: "Frank Ruehl CLM" ,  "Segoe UI" ;
                    font-size: 18px;
                    line-height: 1.5;
                }
            """)
            
            # עדכון ממשק המשתמש למצב תצוגה
            self.edit_button.setText("✍")
            self.edit_button.setToolTip("עריכה")
            self.status_label.setText("מצב תצוגה בלבד")
            self.status_label.setStyleSheet("""
                color: #666666;
                font-size: 14px;
                padding: 5px 15px;
                background-color: transparent;
                border-radius: 10px;
            """)
            
            # הסתרת כפתורי העריכה
            for button in self.editing_buttons:
                button.hide()

        # עדכון מצב הכפתורים
        self.update_buttons_state()

    def on_text_changed(self):
        """מטפל בשינויים בטקסט במצב עריכה"""
        if not self.text_display.isReadOnly():  # רק אם במצב עריכה
            self.save_button.setEnabled(True)        

    def process_text(self, processor_widget):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
            
        try:
            # שמירת התוכן המקורי למקרה של כישלון
            original_content = ""
            if self.current_index >= 0 and self.current_index < len(self.document_history):
                original_content = self.document_history[self.current_index][0]
            

            processor_widget.set_file_path(self.current_file_path)
            self.last_processor_title = processor_widget.windowTitle()
            processor_widget.changes_made.connect(self.refresh_after_processing)
            
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הטקסט: {str(e)}")
            
    def _complete_processing(self, processor_widget, original_content):
        try:
            with open(self.current_file_path, 'r', encoding='utf-8') as file:
                new_content = file.read()

            if new_content != original_content:
                # המרה לתצוגה
                display_content = new_content.replace('\n', '<br>\n')

                self.text_display.setHtml(display_content)

                action_description = self._get_action_description(processor_widget.windowTitle())

                self._safe_update_history(new_content, action_description)
                self.status_label.setText(action_description)

                processor_widget.changes_made.emit()
            else:
                self.status_label.setText("לא בוצעו שינויים")
                
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בהשלמת העיבוד: {str(e)}")
            self.status_label.setText("שגיאה בעיבוד")

    def _get_action_description(self, window_title):
        descriptions = {
            "יצירת כותרות לאוצריא": "בוצעה יצירת כותרות בפורמט אוצריא",
            "יצירת כותרות לאותיות בודדות": "בוצעה יצירת כותרות לאותיות בודדות",
            "הוספת מספר עמוד בכותרת": "נוספו מספרי עמודים בכותרות",
            "שינוי רמת כותרת": "בוצע שינוי ברמת הכותרות",
            "הדגשת מילה ראשונה וניקוד": "בוצעה הדגשת מילים ראשונות והוספת ניקוד",
            "יצירת כותרות לעמוד ב": "נוצרו כותרות לעמוד ב",
            "החלפת כותרות לעמוד ב": "הוחלפו כותרות בעמוד ב",
            "בדיקת שגיאות בכותרות": "בוצעה בדיקת שגיאות בכותרות",
            "בדיקת שגיאות לש\"ס": "בוצעה בדיקת שגיאות מותאמת לש\"ס",
            "המרת תמונה לטקסט": "בוצעה המרת תמונה לטקסט",
            "תיקון שגיאות נפוצות": "בוצע תיקון שגיאות נפוצות",
            "נקודותיים ורווח": "בוצע תיקון נקודותיים ורווחים"
        }
        return descriptions.get(window_title, f"בוצע עיבוד: {window_title}")

    def _safe_update_history(self, content, description):
        try:
            # מחיקת היסטוריה "עתידית" אם קיימת
            if self.current_index < len(self.document_history) - 1:
                self.document_history = self.document_history[:self.current_index + 1]
            
            # הוספת המצב החדש להיסטוריה
            self.document_history.append((content, description))
            self.current_index = len(self.document_history) - 1
            
            # עדכון מצב הכפתורים
            self.update_buttons_state()
            
        except Exception as e:
            print(f"שגיאה בעדכון ההיסטוריה: {str(e)}")


    def undo_action(self):
        """ביטול פעולה אחרונה"""
        try:
            if self.current_index > 0:
                self.current_index -= 1
                content, description = self.document_history[self.current_index]

                display_content = content.replace('\n', '<br>\n')
                self.text_display.setHtml(display_content)

                with open(self.current_file_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                self.status_label.setText(f"בוטל: {description}")
                self.update_buttons_state()
                
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בביטול פעולה: {str(e)}")

    def redo_action(self):
        try:
            if self.current_index < len(self.document_history) - 1:
                self.current_index += 1
                content, description = self.document_history[self.current_index]
                display_content = content.replace('\n', '<br>\n')
                self.text_display.setHtml(display_content)

                with open(self.current_file_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                
                self.status_label.setText(description)
                self.update_buttons_state()
                
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בשחזור פעולה: {str(e)}")

    def update_buttons_state(self):
        try:
            # עדכון כפתורי ביטול וחזרה
            self.undo_button.setEnabled(self.current_index > 0)
            self.redo_button.setEnabled(self.current_index < len(self.document_history) - 1)
            
            # עדכון כפתור שמירה - פעיל אם יש קובץ נוכחי ויש שינויים
            has_changes = len(self.document_history) > 0
            self.save_button.setEnabled(bool(self.current_file_path) and has_changes)
            
            print(f"עדכון מצב כפתורים - קובץ: {bool(self.current_file_path)}, "
                  f"שינויים: {has_changes}, "
                  f"אינדקס: {self.current_index}")
            
        except Exception as e:
            print(f"שגיאה בעדכון מצב הכפתורים: {str(e)}")
            
    def save_file(self):
        if not self.current_file_path:
            self.save_file_as()
            return
            
        try:
            if self.current_index >= 0 and self.current_index < len(self.document_history):
                # שימוש בתוכן המקורי מההיסטוריה
                content = self.document_history[self.current_index][0]
            else:
                content = self.text_display.toHtml()
                # הסרת כל התגים של HTML
                content = re.sub(r'<[^>]+>', '', content)
                content = content.replace('&nbsp;', ' ')  # החלפת רווחים מיוחדים
            
            with open(self.current_file_path, 'w', encoding='utf-8') as file:
                file.write(content)
                
            self.status_label.setText("הקובץ נשמר בהצלחה")
            QMessageBox.information(self, "שמירה", "הקובץ נשמר בהצלחה!")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")

    def save_file(self):
        """שמירת הקובץ"""
        try:
            if not self.current_file_path:
                return
            current_content = self.document_history[self.current_index][0]
            with open(self.current_file_path, 'w', encoding='utf-8') as file:
                file.write(current_content)
            
            self.status_label.setText("הקובץ נשמר בהצלחה")
            
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")
            
    def select_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "בחר קובץ טקסט",
            "",
            "קבצי טקסט (*.txt)",
            
            options=options
        )
        
        if file_path:
            self.current_file_path = file_path
            self.load_file(file_path)

    def load_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                
            # המרת רווחים וירידות שורה לתגים מתאימים תוך שמירה על תגי HTML קיימים
            processed_content = content.replace('\n', '<br>\n')  # מוסיף <br> לפני כל ירידת שורה
            
            self.text_display.setHtml(processed_content)
            self.document_history = [(content, "מצב התחלתי")]  # שומר את התוכן המקורי ללא תגי BR
            self.current_index = 0
            self.update_buttons_state()
            self.status_label.setText("קובץ נטען בהצלחה")
            
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בטעינת הקובץ: {str(e)}")


    def save_action(self):
        if not self.current_file_path:
            self.save_file_as()
            return
            
        try:
            content = self.text_display.toHtml()
            # המרה חזרה - הסרת תגי <br> 
            content = content.replace('<br>\n', '\n')
            
            with open(self.current_file_path, 'w', encoding='utf-8') as file:
                file.write(content)
                
            self.status_label.setText("הקובץ נשמר בהצלחה")
            QMessageBox.information(self, "שמירה", "הקובץ נשמר בהצלחה!")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")



    def refresh_after_processing(self):
        """עדכון התצוגה לאחר עיבוד"""
        try:
            with open(self.current_file_path, 'r', encoding='utf-8') as file:
                new_content = file.read()

            display_content = new_content.replace('\n', '<br>\n')
            self.text_display.setHtml(display_content)

            if self.last_processor_title:
                action_description = self._get_action_description(self.last_processor_title)
                # עדכון ההיסטוריה והסטטוס
                self._safe_update_history(new_content, action_description)
                self.status_label.setText(action_description)
                
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעדכון התצוגה: {str(e)}")

    def _get_action_description(self, window_title):
        """קבלת תיאור מפורט של הפעולה לפי שם החלון"""
        descriptions = {
            "יצירת כותרות לאוצריא": "בוצעה יצירת כותרות בפורמט אוצריא",
            "יצירת כותרות לאותיות בודדות": "בוצעה יצירת כותרות לאותיות בודדות",
            "הוספת מספר עמוד בכותרת": "נוספו מספרי עמודים בכותרות",
            "שינוי רמת כותרת": "בוצע שינוי ברמת הכותרות",
            "הדגשת מילה ראשונה וניקוד": "בוצעה הדגשת מילים ראשונות והוספת ניקוד",
            "יצירת כותרות לעמוד ב": "נוצרו כותרות לעמוד ב",
            "החלפת כותרות לעמוד ב": "הוחלפו כותרות בעמוד ב",
            "בדיקת שגיאות בכותרות": "בוצעה בדיקת שגיאות בכותרות",
            "בדיקת שגיאות לש\"ס": "בוצעה בדיקת שגיאות מותאמת לש\"ס",
            "המרת תמונה לטקסט": "בוצעה המרת תמונה לטקסט",
            "תיקון שגיאות נפוצות": "בוצע תיקון שגיאות נפוצות",
            "נקודותיים ורווח": "בוצע תיקון נקודותיים ורווחים",
            "סקריפט 1": "בוצע עיבוד סקריפט 1",  # הוספת תיאורים לסקריפטים
            "סקריפט 2": "בוצע עיבוד סקריפט 2",
            
        }
        return descriptions.get(window_title, f"בוצע עיבוד: {window_title}")
   
    def update_content_from_child(self):
        """עדכון התצוגה לאחר שינויים בחלונות המשנה"""
        if not self.current_file_path:
            return
            
        try:
            with open(self.current_file_path, 'r', encoding='utf-8') as file:
                new_content = file.read()
            
            # המרה לתצוגה
            display_content = new_content.replace('\n', '<br>\n')
            self.text_display.setHtml(display_content)

            action_description = self._get_action_description(self.last_processor_title)
            self._safe_update_history(new_content, action_description)

            print(f"עודכן תוכן חדש - תיאור: {action_description}")
            print(f"גודל היסטוריה: {len(self.document_history)}")
            
        except Exception as e:
            print(f"שגיאה בעדכון התוכן: {str(e)}")
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעדכון התוכן: {str(e)}")

 #סנכרון חלונות המשנה עם החלון הראשי

    def open_about_dialog(self):
        """פתיחת חלון 'אודות'"""
        dialog = AboutDialog(self)
        dialog.exec_()

            
    # סקריפט 1 - יצירת כותרות לאוצריא
    def open_create_headers_otzria(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.create_headers_window = CreateHeadersOtZria()
        self.create_headers_window.set_file_path(self.current_file_path)
        self.create_headers_window.changes_made.connect(self.refresh_after_processing)
        self.create_headers_window.show()

    # סקריפט 2 - יצירת כותרות לאותיות בודדות
    def open_create_single_letter_headers(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.create_single_letter_headers_window = CreateSingleLetterHeaders()
        self.create_single_letter_headers_window.set_file_path(self.current_file_path)
        self.create_single_letter_headers_window.changes_made.connect(self.update_content_from_child)
        self.create_single_letter_headers_window.show()

    # סקריפט 3 - הוספת מספר עמוד בכותרת הדף
    def open_add_page_number_to_heading(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.add_page_number_window = AddPageNumberToHeading()
        self.add_page_number_window.set_file_path(self.current_file_path)
        self.add_page_number_window.changes_made.connect(self.update_content_from_child)
        self.add_page_number_window.show()

    # סקריפט 4 - שינוי רמת כותרת
    def open_change_heading_level(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.change_heading_level_window = ChangeHeadingLevel()
        self.change_heading_level_window.set_file_path(self.current_file_path)
        self.change_heading_level_window.changes_made.connect(self.update_content_from_child)
        self.change_heading_level_window.show()

    # סקריפט 5 - הדגשת מילה ראשונה וניקוד בסוף קטע
    def open_emphasize_and_punctuate(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.emphasize_and_punctuate_window = EmphasizeAndPunctuate()
        self.emphasize_and_punctuate_window.set_file_path(self.current_file_path)
        self.emphasize_and_punctuate_window.changes_made.connect(self.update_content_from_child)
        self.emphasize_and_punctuate_window.show()

    # סקריפט 6 - יצירת כותרות לעמוד ב
    def open_create_page_b_headers(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.create_page_b_headers_window = CreatePageBHeaders()
        self.create_page_b_headers_window.set_file_path(self.current_file_path)
        self.create_page_b_headers_window.changes_made.connect(self.update_content_from_child)
        self.create_page_b_headers_window.show()

    # סקריפט 7 - החלפת כותרות לעמוד ב
    def open_replace_page_b_headers(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.replace_page_b_headers_window = ReplacePageBHeaders()
        self.replace_page_b_headers_window.set_file_path(self.current_file_path)
        self.replace_page_b_headers_window.changes_made.connect(self.update_content_from_child)
        self.replace_page_b_headers_window.show()

    # סקריפט 8 - בדיקת שגיאות בכותרות
    def open_check_heading_errors_original(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.check_heading_errors_original_window = CheckHeadingErrorsOriginal()
        self.check_heading_errors_original_window.set_file_path(self.current_file_path)
        self.check_heading_errors_original_window.process_file(self.current_file_path)
        self.check_heading_errors_original_window.show()

    # סקריפט 9 - בדיקת שגיאות בכותרות מותאם לספרים על הש"ס
    def open_check_heading_errors_custom(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.check_heading_errors_custom_window = CheckHeadingErrorsCustom()
        self.check_heading_errors_custom_window.set_file_path(self.current_file_path)
        self.check_heading_errors_custom_window.process_file(self.current_file_path)
        self.check_heading_errors_custom_window.show()

    # סקריפט 10 - המרת תמונה לטקסט
    def open_Image_To_Html_App(self):
        self.Image_To_Html_App_window = ImageToHtmlApp()
        self.Image_To_Html_App_window.show()

    # סקריפט 11 - תיקון שגיאות נפוצות
    def open_Text_Cleaner_App(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "שגיאה", "נא לבחור קובץ תחילה")
            return
        self.Text_Cleaner_App_window = TextCleanerApp()
        self.Text_Cleaner_App_window.set_file_path(self.current_file_path)
        self.Text_Cleaner_App_window.changes_made.connect(self.update_content_from_child)
        self.Text_Cleaner_App_window.show()

   
        

    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
    
    #עדכונים
    def check_for_updates(self):
        
        """בדיקת עדכונים חדשים"""
        self.status_label.setText("בודק עדכונים...")

        self.update_checker = UpdateChecker(self.current_version)

        self.update_checker.update_available.connect(self.handle_update_available)
        self.update_checker.no_update.connect(self.handle_no_update)
        self.update_checker.error.connect(self.handle_update_error)

        self.update_checker.start()

    def handle_update_available(self, download_url, new_version):
        """טיפול בעידכון"""
        reply = QMessageBox.question(
            self,
            "נמצא עדכון",
            f"נמצאה גרסה חדשה ({new_version})!\nהאם ברצונך לעדכן כעת?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.download_and_install_update(download_url, new_version)
        else:
            self.status_label.setText("עדכון זמין")

    def handle_no_update(self):
        """טיפול במקרה שאין עדכון"""
        QMessageBox.information(
            self,
            "אין עדכונים",
            "התוכנה מעודכנת לגרסה האחרונה"
        )
        self.status_label.setText("התוכנה מעודכנת")

    def handle_update_error(self, error_msg):
        """טיפול בשגיאות בתהליך העדכון"""
        QMessageBox.warning(
            self,
            "שגיאה",
            error_msg
        )
        self.status_label.setText("שגיאה בבדיקת עדכונים")

    def download_and_install_update(self, download_url, new_version):
        """הורדת והתקנת העדכון"""
        try:
            self.status_label.setText("מוריד עדכון...")
            response = requests.get(download_url, stream=True)
            response.raise_for_status()

            current_exe = sys.executable
            backup_exe = current_exe + '.backup'
            new_exe = current_exe + '.new'

            with open(new_exe, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
             #יצירת קובץ bat לטיפול בעידכון
            update_script = f"""
            @echo off
            timeout /t 1 /nobreak > nul
            move "{current_exe}" "{backup_exe}"
            move "{new_exe}" "{current_exe}"
            start "" "{current_exe}"
            del "%~f0"
            """.strip()

            update_bat = os.path.join(os.path.dirname(current_exe), 'update.bat')
            with open(update_bat, 'w') as f:
                f.write(update_script)

            QMessageBox.information(
                self,
                "התקנת עדכון",
                "העדכון ירד בהצלחה. התוכנה תיסגר כעת ותופעל מחדש עם הגרסה החדשה."
            )
            os.startfile(update_bat)
            sys.exit()

        except Exception as e:
            QMessageBox.critical(
                self,
                "שגיאה",
                f"שגיאה בהורדת העדכון: {str(e)}"
            )
            self.status_label.setText("שגיאה בהורדת העדכון")      


class AboutDialog(QDialog):
    """חלון 'אודות'"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("אודות התוכנה")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFixedWidth(600)
        self.setStyleSheet("""
            QDialog {
                background-color: white;
                border-radius: 15px;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # כותרת ראשית
        title_label = QLabel("עריכת ספרי דיקטה עבור 'אוצריא'")
        title_label.setStyleSheet("""
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 24px;
                font-weight: bold;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # מידע על הגרסה והתאריך
        info_container = QHBoxLayout()
        
        version_label = QLabel("גירסה: v3.2")
        version_label.setStyleSheet("""
            QLabel {
                color: #2b4c7e;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        
        date_label = QLabel("תאריך: כט שבט תשפה")
        date_label.setStyleSheet(version_label.styleSheet())
        
        info_container.addWidget(version_label, alignment=Qt.AlignCenter)
        info_container.addWidget(date_label, alignment=Qt.AlignCenter)
        layout.addLayout(info_container)

        # פיתוח
        dev_label = QLabel("נכתב על ידי 'מתנדבי אוצריא', להצלחת לומדי התורה הקדושה")
        dev_label.setStyleSheet("""
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin: 10px 0;
            }
        """)
        dev_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(dev_label)

        # סגנון משותף לקישורים
        link_style = """
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                padding: 5px;
            }
            QLabel a {
                color: #2b4c7e;
                text-decoration: none;
            }
            QLabel a:hover {
                color: #1a365d;
                text-decoration: underline;
            }
        """

        # קישורים להורדה
        github_label = QLabel('ניתן להוריד את הגירסא האחרונה, וכן קובץ הדרכה, בקישור הבא: <a href="https://github.com/YOSEFTT/EditingDictaBooks/releases">GitHub</a>')
        github_label.setStyleSheet(link_style)
        github_label.setOpenExternalLinks(True)
        github_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(github_label)

        mitmachimtop_label = QLabel('או כאן: <a href="https://mitmachim.top/topic/77509/הסבר-הוספת-וטיפול-בספרים-ל-אוצריא-כעת-זה-קל">מתמחים.טופ</a>')
        mitmachimtop_label.setStyleSheet(link_style)
        mitmachimtop_label.setOpenExternalLinks(True)
        mitmachimtop_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(mitmachimtop_label)

        drive_label = QLabel('או בדרייב: <a href="http://did.li/otzaria-">כאן</a> או <a href="https://drive.google.com/open?id=1KEKudpCJUiK6Y0Eg44PD6cmbRsee1nRO&usp=drive_fs">כאן</a>')
        drive_label.setStyleSheet(link_style)
        drive_label.setOpenExternalLinks(True)
        drive_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(drive_label)

        # מידע נוסף
        info_text = "אפשר לבקש את התוכנה\nוכן להירשם לקבלת עדכון במייל כשיוצא עדכון לתוכנות אלו\nוכן לקבל תמיכה וסיוע בכל הקשור לתוכנה זו ולתוכנת 'אוצריא'\nבמייל הבא:"
        info_label = QLabel(info_text)
        info_label.setStyleSheet("""
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 14px;
                margin: 15px 0 5px 0;
            }
        """)
        info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(info_label)

        # כתובת מייל
        gmail_label = QLabel('<a href="https://mail.google.com/mail/u/0/?view=cm&fs=1&to=otzaria.1%40gmail.com%E2%80%AC">otzaria.1@gmail.com</a>')
        gmail_label.setStyleSheet("""
            QLabel {
                color: #2b4c7e;
                font-family: "Segoe UI", Arial;
                font-size: 16px;
                font-weight: bold;
            }
            QLabel a {
                color: #2b4c7e;
                text-decoration: none;
            }
            QLabel a:hover {
                color: #1a365d;
                text-decoration: underline;
            }
        """)
        gmail_label.setOpenExternalLinks(True)
        gmail_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(gmail_label)

        self.setLayout(layout)

        
class DocumentHistory:
    def __init__(self, max_stack_size=100):
        self.undo_stack = []  
        self.redo_stack = []
        self.max_stack_size = max_stack_size
        self.current_content = ""
        self.current_description = "לא בוצעו עדיין פעולות"

    def push_state(self, content, description):

        if content != self.current_content:
            self.undo_stack.append((self.current_content, self.current_description))
            self.current_content = content
            self.current_description = description
            self.redo_stack.clear()

            if len(self.undo_stack) > self.max_stack_size:
                self.undo_stack.pop(0)

    def undo(self):
        """ביטול פעולה אחרונה"""
        if not self.can_undo():
            return None, None

        self.redo_stack.append((self.current_content, self.current_description))
        self.current_content, self.current_description = self.undo_stack.pop()
        return self.current_content, self.current_description

    def redo(self):
        """חזרה על פעולה שבוטלה"""
        if not self.can_redo():
            return None, None

        self.undo_stack.append((self.current_content, self.current_description))
        self.current_content, self.current_description = self.redo_stack.pop()
        return self.current_content, self.current_description

    def get_current_description(self):
        """קבלת תיאור הפעולה הנוכחית"""
        return self.current_description

 
    
# ==========================================
#  update
# ==========================================
class UpdateChecker(QThread):
    """בדיקת עידכונים"""
    update_available = pyqtSignal(str, str)  
    no_update = pyqtSignal()  
    error = pyqtSignal(str)  

    def __init__(self, current_version):
        super().__init__()
        self.current_version = current_version
        
    def run(self):
        try:
            response = requests.get(
                "https://github.com/YOSEFTT/EditingDictaBooks/releases/latest",
                timeout=10
            )
            response.raise_for_status()
            
            latest_release = response.json()
            latest_version = latest_release['tag_name'].lstrip('v')
            
            if version.parse(latest_version) > version.parse(self.current_version):
                # מצאנו גרסה חדשה יותר
                download_url = None
                for asset in latest_release['assets']:
                    if asset['name'].endswith('.exe'):  
                        download_url = asset['browser_download_url']
                        break
                
                if download_url:
                    self.update_available.emit(download_url, latest_version)
                else:
                    self.error.emit("נמצאה גרסה חדשה אך לא נמצא קובץ הורדה מתאים")
            else:
                self.no_update.emit()
        
        except Exception as e:
            self.error.emit(f"שגיאה בבדיקת עדכונים: {str(e)}")

   
# ==========================================
# Main Application
# ==========================================
def main():
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    main_menu = MainMenu()
    main_menu.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
