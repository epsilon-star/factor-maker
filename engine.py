import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import docxtpl
from docx2pdf import convert
import os
import shutil

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(APP_ROOT, "output")
IMAGES_DIR = os.path.join(APP_ROOT, "images")

persian_numbers = {
	0: "صفر", 1: "یک", 2: "دو", 3: "سه", 4: "چهار", 5: "پنج", 6: "شش", 7: "هفت", 8: "هشت", 9: "نه",
	10: "ده", 11: "یازده", 12: "دوازده", 13: "سیزده", 14: "چهارده", 15: "پانزده", 16: "شانزده", 17: "هفده", 18: "هجده", 19: "نوزده",
	20: "بیست", 30: "سی", 40: "چهل", 50: "پنجاه", 60: "شصت", 70: "هفتاد", 80: "هشتاد", 90: "نود",
	100: "صد", 200: "دویست", 300: "سیصد", 400: "چهارصد", 500: "پانصد", 600: "ششصد", 700: "هفتصد", 800: "هشتصد", 900: "نهصد"
}

powers_of_thousand = [
	(1000000000, "میلیارد"),  # Billion
	(1000000, "میلیون"),      # Million
	(1000, "هزار"),           # Thousand
]

def_params = {
	"factor_id":"",
	"factor_date":"",
	"factor_addons":"",

	"title":"",
	"sby":"",
	"phone":"",
	"location":"",

	"factor_price_name":"",
	"factor_price_number":"",
	"factor_disc_1":"",
	"factor_disc_2":"",
}

image_map = {
	"blue": "template-1",
	"gray": "template-gray",
	"red": "template-3",
	"green": "template-green",
}

DEFAULT_THEMES = ["blue", "gray"]
THEME_LABELS = {
	"blue": "آبی",
	"gray": "خاکستری",
	"red": "قرمز",
	"green": "سبز",
}

class PersianFactor:
	def __init__(self,msgShow = None):
		self.msgShow = msgShow
		pass

	def makeFactor(self, ids: int, template_path: str, datas: dict, themes: list | None = None) -> bool:
		"""Make an invoice using templates for the selected color themes."""

		try:
			selected = themes if themes else list(DEFAULT_THEMES)
			invalid = [t for t in selected if t not in image_map]
			if invalid:
				raise ValueError(f"تم ناشناخته: {', '.join(invalid)}")
			if not selected:
				raise ValueError("حداقل یک رنگ باید انتخاب شود")

			buffer = def_params.copy()
			buffer.update(datas)
			totalPrice = self.calculateTotal(list(buffer['items'].values()), datas['tax'], datas['off'])  # type: ignore
			buffer['factor_price_name'] = self.convertWord(totalPrice) + " ریال"
			buffer['factor_price_number'] = self.formatNumber(str(totalPrice))

			# PHASE 1
			doc = docx.Document(template_path)
			dtab = doc.tables[0]
			idx = 0
			for x in buffer['items']:
				buffer['items'][x]  # type: ignore
				res = dtab.cell(idx, 3)
				res.text = x
				res.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
				res.paragraphs[0].style.font.size = Pt(7)  # type: ignore
				res = dtab.cell(idx, 2)
				res.text = self.persianNumber(str(buffer['items'][x][0]))  # type: ignore
				res.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
				res.paragraphs[0].style.font.size = Pt(7)  # type: ignore
				res = dtab.cell(idx, 1)
				res.text = self.formatNumber(str(buffer['items'][x][1]))  # type: ignore
				res.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
				res.paragraphs[0].style.font.size = Pt(7)  # type: ignore
				res = dtab.cell(idx, 0)
				res.text = self.formatNumber(str(round(buffer['items'][x][1] * buffer['items'][x][0])))  # type: ignore
				res.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
				res.paragraphs[0].style.font.size = Pt(7)  # type: ignore
				idx += 1
			for x in doc.tables:
				x.style.font.name = "Vazir"  # type: ignore

			idx_dir = os.path.join(OUTPUT_DIR, str(ids))
			os.makedirs(idx_dir, exist_ok=True)
			base_path = os.path.join(idx_dir, f"{ids}-base.docx")
			doc.save(base_path)

			generated = []
			for theme_name in selected:
				image_name = image_map[theme_name]
				output_path = os.path.join(idx_dir, f"{ids}-{theme_name}.docx")
				image_path = os.path.join(IMAGES_DIR, f"{image_name}.jpg")
				shutil.copyfile(base_path, output_path)
				tpl = docxtpl.DocxTemplate(output_path)
				tpl.replace_pic("background_image", image_path)
				tpl.render(buffer)
				tpl.save(output_path)
				self.getPDF(output_path)
				generated.append(theme_name)

			try:
				os.remove(base_path)
			except OSError:
				pass

			theme_names = "، ".join(THEME_LABELS.get(t, t) for t in generated)
			if self.msgShow:
				self.msgShow(f"فاکتور با آیدی [{ids}] ساخته شد\nرنگ‌ها: {theme_names}")
			return True
		except Exception as EXP:
			if self.msgShow:
				self.msgShow(f"خطا در موتور\nPersianFactor.makeFactor() -> [{EXP}]")
			return False

	def calculateTotal(self,items:list,tax:float,off:float):
		""" Calculate The Total Price In The END """
		try:	
			buffer = items.copy()
			buff = 0
			for x in buffer:
				buff += x[0] * x[1]
			buff *= (100 + (tax))/100
			buff *= (100 - (off))/100
			return round(buff)
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.calculateTotal() -> [{EXP}]")
			return 0.0

	def getPDF(self,path:str):
		""" Convert DocX Files To PDF """
		try:
			return convert(path)
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.getPDF() -> [{EXP}]")

	def convertHundred(self,number) -> str:
		""" Convert numbers from 1 to 999 to Persian words """
		try:
			if number < 20:
				return persian_numbers[number]
			elif number < 100:
				tens = number // 10 * 10
				remainder = number % 10
				return persian_numbers[tens] + ("" if remainder == 0 else " و " + persian_numbers[remainder])
			else:
				hundreds = number // 100 * 100
				remainder = number % 100
				return persian_numbers[hundreds] + ("" if remainder == 0 else " و " + self.convertHundred(remainder))
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.convertHundred() -> [{EXP}]")
			return ''

	def convertWord(self,number) -> str:
		""" Convert any number to Persian words """
		try:
			if number == 0:
				return persian_numbers[0]

			words = []

			# Iterate over powers of thousand
			for value, word in powers_of_thousand:
				if number >= value:
					count = number // value
					words.append(self.convertHundred(count) + " " + word)
					number %= value

			if number > 0:
				words.append(self.convertHundred(number))

			return " و ".join(words)
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.convertWord() -> [{EXP}]")
			return ''

	def persianNumber(self,number:str) -> str:
		""" Convert Any English Number To Persian Format """
		try:
			perDigits = {
				'0': '۰', '1': '۱', '2': '۲', '3': '۳', '4': '۴', 
				'5': '۵', '6': '۶', '7': '۷', '8': '۸', '9': '۹'
			}
			return ''.join(perDigits[x] if x in perDigits else x for x in str(number))
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.persianNumber() -> [{EXP}]")
			return ''

	def formatNumber(self,number:str) -> str:
		""" Gives ، To Every 3 Numbers Arranged """
		try:
			output = str(self.persianNumber(number))[::-1]
			buffer = [output[x:x+3] for x in range(0,len(output),3)]
			return '،'.join(buffer)[::-1]
		except Exception as EXP:
			if self.msgShow: self.msgShow(f"خطا در موتور\nPersianFactor.formatNumber() -> [{EXP}]")
			return ''

if __name__ == "__main__":
	d = PersianFactor()
	ids = 40401004
	# ids = 157565
	lts = {
		# "نصب و راه‌اندازی نرم‌افزار آریا":[11,50000],
		# "نصب و راه‌اندازی نرم‌افزار میشا و کوشا":[2,70000],
		# "فعال‌سازی و کرک نرم‌افزار آریا":[1,1300000],
		# "فعال‌سازی و کرک نرم‌افزار میشا و کوشا":[1,400000],
		# "بازیابی و ریکاوری سیستم":[1,250000],
		# "پاکسازی ویروس و بهینه‌سازی سیستم":[3,100000],
		# "فعال‌سازی پرینتر و راه‌اندازی اولیه":[1,100000]
		# "نصب آنتی ویروس":[1,1500000],
		# "نصب و راه اندازی ویندوز":[1,1500000],
		# "نصب درایورها و نرم‌افزارها و آنتی ویروس":[1,1500000],
		# "مونتاژ و نصب کیس سیستم":[1,2000000],
		# "کیس سیستم":[1,7500000],
		# "ارتقای رم 4 گیگابایت":[2,5000000],
		# "موس":[1,3000000],
		"هارد اینترنال 1 ترابایت وسترن آبی":[1,35000000],
		"هارد اینترنال 2 ترابایت وسترن بنفش":[1,68000000],
		"نصب هاردیسک":[1,2000000],
		# "":[1,0],
		# "سوئیچ 16 کاناله + 2 آپلینک":[1,60000000],
		# "دستگاه NVR داهوا 16 کاناله شبکه":[1,60000000],
		# "سوکت 30 عددی Cat6 PTC":[1,1000000],
		# "کابل شبکه Cat6 PVC 305 متری":[300,135000],
		# "دوربین مداربسته داهوا شبکه میکروفون":[8,35000000],
		# "نصب سویچ + نصب NVR":[1,9000000],
		# "کابل کشی متری":[300,150000],
		# "نصب دوربین مداربسته":[8,3500000],
		# "":[1,0],
	}
	datas = {
		"factor_id":d.persianNumber(str(ids)),
		"factor_date":d.persianNumber("1404/01/30"),
		"factor_addons":"",

		"title":"فاکتور فروش",
		"sby":"تیم توسعه دهنده متانویا",
		"phone":"",
		"location":"تهران، منطقه 10، بلوار ستارخان، خیابان توکلی، انتهای کوچه خامنه، دبیرستان دخترانه ولایت فقیه",

		"items":lts,
		"off":0.0,
		"tax":0.0,

		# "factor_disc_1":d.persianNumber("6219 - 8618 - 2621 - 0253\n6219 - 8619 - 8779 - 8203"),
		"factor_disc_1":d.persianNumber("5859 - 8311 - 6428 - 7749"),
		"factor_disc_2":"مریم رفیعی",
		# "factor_disc_2":"مبین براتیان تاج الدینی - مانی حامی کمری",
	}
	d.makeFactor(ids,"templates/template.docx",datas)