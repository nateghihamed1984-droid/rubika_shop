from rubka import Robot, Message
from rubka.keypad import ChatKeypadBuilder
import sqlite3
import os
from datetime import datetime, timedelta

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

TOKEN = "CFFHIC0ZOTGPALAPFBHEWCNUDYSNDTKTENGOAULXCZPLZRSBYTHFVPMKUWLMEPHM"
ADMIN_CHAT_ID = "b0BC4FX0BHkJ080af57b490e17163d49"
DB_PATH = "shop.db"
PRODUCT_CODE = "P-001"
PRODUCT_IMAGE_PATH = "product_p001.jpg"
FREE_DELIVERY_CITY = "تبریز"
RECENT_ORDER_DAYS = 7
RECENT_ORDER_THRESHOLD = 3
RETURN_WINDOW_HOURS = 24
EXCHANGE_WINDOW_HOURS = 72

DEFAULT_PRODUCT = {
    "code": PRODUCT_CODE,
    "name": "محصول نمونه",
    "description": "توضیحات محصول را اینجا وارد کنید.",
    "price": "1500000",
    "delivery_days": "1 تا 3",
    "image_path": PRODUCT_IMAGE_PATH,
    "active": 1,
}

conn = sqlite3.connect(DB_PATH, check_same_thread=False)
cursor = conn.cursor()

def cols(table):
    cursor.execute(f"PRAGMA table_info({table})")
    return [r[1] for r in cursor.fetchall()]

def addcol(table, name, definition):
    if name not in cols(table):
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")
        conn.commit()

cursor.execute("""CREATE TABLE IF NOT EXISTS orders(
 id INTEGER PRIMARY KEY AUTOINCREMENT,user_id TEXT,product TEXT,name TEXT,phone TEXT,
 address TEXT,status TEXT)""")
for n,d in [
 ("chat_id","TEXT"),("created_at","TEXT"),("delivered_at","TEXT"),
 ("accepted_at","TEXT"),("customer_acceptance","TEXT")]: addcol("orders",n,d)

cursor.execute("""CREATE TABLE IF NOT EXISTS products(
 id INTEGER PRIMARY KEY AUTOINCREMENT,code TEXT UNIQUE,name TEXT,description TEXT,
 price TEXT,delivery_days TEXT,image_path TEXT,active INTEGER DEFAULT 1)""")
addcol("products","image_file_id","TEXT")

cursor.execute("""CREATE TABLE IF NOT EXISTS reviews(
 id INTEGER PRIMARY KEY AUTOINCREMENT,order_id INTEGER,user_id TEXT,product TEXT,
 rating INTEGER,comment TEXT,created_at TEXT)""")

cursor.execute("""CREATE TABLE IF NOT EXISTS return_requests(
 id INTEGER PRIMARY KEY AUTOINCREMENT,order_id INTEGER UNIQUE,user_id TEXT,product TEXT,
 request_type TEXT,reason TEXT,status TEXT,created_at TEXT,decided_at TEXT,admin_note TEXT)""")

# New after-sales fields are added without destroying existing databases.
for n,d in [
 ("return_deadline","TEXT"),("customer_name","TEXT"),("phone","TEXT"),
 ("exchange_product","TEXT"),("exchange_difference","TEXT"),
 ("refund_amount","TEXT"),("bank_account_name","TEXT"),("card_number","TEXT"),
 ("iban","TEXT"),("bank_name","TEXT"),("paid_at","TEXT"),("exchange_delivery_days","TEXT"),("exchange_sent_at","TEXT"),("exchange_received_at","TEXT"),("exchange_customer_acceptance","TEXT"),("exchange_round","INTEGER DEFAULT 1")]:
    addcol("return_requests",n,d)

cursor.execute("SELECT id FROM products WHERE code=?",(PRODUCT_CODE,))
if not cursor.fetchone():
    cursor.execute("""INSERT INTO products(code,name,description,price,delivery_days,image_path,active)
                     VALUES(?,?,?,?,?,?,?)""", tuple(DEFAULT_PRODUCT.values()))
    conn.commit()

bot = Robot(token=TOKEN)
customers = {}
rating_comments = {}
admin_states = {}
after_sales_states = {}


def now_string(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def parse_dt(v):
    try: return datetime.strptime(v,"%Y-%m-%d %H:%M:%S") if v else None
    except: return None
def get_user_id(m): return getattr(m,"sender_id",None)
def get_chat_id(m): return getattr(m,"chat_id",None)
def is_admin(m): return get_chat_id(m)==ADMIN_CHAT_ID
def button_id(m):
    try: return m.aux_data.button_id
    except: return None

def norm(v):
    return str(v or "").translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩","01234567890123456789"))
def price(v):
    try: return f"{int(norm(v)):,}"
    except: return str(v or "نامشخص")
def product():
    cursor.execute("SELECT code,name,description,price,delivery_days,image_path,image_file_id FROM products WHERE code=? AND active=1",(PRODUCT_CODE,))
    return cursor.fetchone()
def admin_product():
    cursor.execute("SELECT code,name,description,price,delivery_days,image_path,image_file_id,active FROM products WHERE code=?",(PRODUCT_CODE,))
    return cursor.fetchone()
def deadline(accepted, hours=RETURN_WINDOW_HOURS):
    d=parse_dt(accepted); return d+timedelta(hours=hours) if d else None

def status_fa(s):
    return {"جدید":"🆕 جدید","تأیید شده":"✅ تأیید شده","لغو شده":"❌ لغو شده","تحویل شده":"🚚 تحویل شده","تحویل و تأیید شده":"📦 تحویل و تأیید شده","تعویض انجام شد":"🔄 تعویض انجام شد","مرجوع و وجه بازگشت داده شد":"💰 مرجوع و وجه بازگشت داده شد"}.get(s,s or "نامشخص")
def after_status(s):
    return {"در انتظار بررسی":"🟡 در انتظار بررسی","تأیید شده":"✅ تأیید شده","رد شده":"❌ رد شده","در حال تعویض":"🔄 در حال تعویض","تعویض انجام شد":"✅ تعویض انجام شد","کالای تعویضی ارسال شد":"📦 کالای تعویضی ارسال شد","در انتظار تأیید دریافت تعویضی":"🕐 در انتظار تأیید دریافت تعویضی","در انتظار بررسی مجدد":"🟡 در انتظار بررسی مجدد","در انتظار اطلاعات بانکی":"🏦 در انتظار اطلاعات بانکی","در انتظار واریز":"💰 در انتظار واریز","وجه واریز شد":"✅ وجه واریز شد","تکمیل شده":"✅ تکمیل شده"}.get(s,s or "نامشخص")

def main_menu():
    return (ChatKeypadBuilder().row(ChatKeypadBuilder().button(id="show_products",text="📦 مشاهده محصولات"))
      .row(ChatKeypadBuilder().button(id="my_orders",text="📋 مشاهده سفارش‌ها"))
      .row(ChatKeypadBuilder().button(id="shop_policy",text="📜 شرایط تعویض و مرجوعی")).build())
def product_menu():
    return (ChatKeypadBuilder().row(ChatKeypadBuilder().button(id="order_product_P-001",text="🛒 ثبت سفارش"))
      .row(ChatKeypadBuilder().button(id="main_menu",text="🔙 منوی اصلی")).build())
def confirm_menu():
    return ChatKeypadBuilder().row(ChatKeypadBuilder().button(id="customer_confirm_order",text="✅ تأیید سفارش"),ChatKeypadBuilder().button(id="customer_cancel_order",text="❌ لغو سفارش")).build()
def accept_menu(oid):
    return ChatKeypadBuilder().row(ChatKeypadBuilder().button(id=f"customer_accept_delivery_{oid}",text="✅ کالا را رویت کردم و سالم و مطابق سفارش تحویل گرفتم")).build()
def order_admin_menu(oid):
    return (ChatKeypadBuilder().row(ChatKeypadBuilder().button(id=f"admin_confirm_{oid}",text="✅ تأیید سفارش"),ChatKeypadBuilder().button(id=f"admin_cancel_{oid}",text="❌ لغو سفارش"))
      .row(ChatKeypadBuilder().button(id=f"admin_deliver_{oid}",text="🚚 تحویل شد")).build())
def return_admin_menu(rid,typ,status="در انتظار بررسی"):
    b=ChatKeypadBuilder()
    if status in ("در انتظار بررسی","در انتظار بررسی مجدد"):
        return b.row(ChatKeypadBuilder().button(id=f"admin_after_approve_{rid}",text="✅ تأیید درخواست"),ChatKeypadBuilder().button(id=f"admin_after_reject_{rid}",text="❌ رد درخواست")).build()
    if typ=="مرجوعی و بازگشت وجه" and status=="در انتظار واریز":
        return b.row(ChatKeypadBuilder().button(id=f"admin_refund_paid_{rid}",text="💰 ثبت واریز وجه")).build()
    if typ=="تعویض" and status=="در حال تعویض":
        return b.row(ChatKeypadBuilder().button(id=f"admin_exchange_sent_{rid}",text="📦 ارسال کالای تعویضی")).build()
    if typ=="تعویض" and status=="در انتظار تأیید دریافت تعویضی":
        return b.row(ChatKeypadBuilder().button(id=f"admin_exchange_done_{rid}",text="🔄 ثبت نهایی تعویض")).build()
    return b.build()
def customer_after_menu(oid, exchange=False):
    # منوی تصمیم پس از تحویل اولیه یا تحویل کالای تعویضی
    title = "⭐ راضی هستم؛ ثبت امتیاز و نظر"
    return (ChatKeypadBuilder().row(
                ChatKeypadBuilder().button(id=f"customer_rate_{oid}",text=title)
            )
            .row(
                ChatKeypadBuilder().button(id=f"customer_exchange_{oid}",text="🔄 مجدداً درخواست تعویض" if exchange else "🔄 درخواست تعویض"),
                ChatKeypadBuilder().button(id=f"customer_refund_{oid}",text="💰 مرجوعی و بازگشت وجه")
            ).build())

def exchange_accept_menu(rid):
    return ChatKeypadBuilder().row(
        ChatKeypadBuilder().button(id=f"customer_accept_exchange_{rid}",text="✅ کالای تعویضی را دریافت کردم")
    ).build()

def refund_accept_menu(rid):
    return ChatKeypadBuilder().row(
        ChatKeypadBuilder().button(id=f"customer_accept_refund_{rid}",text="✅ دریافت وجه را تأیید می‌کنم")
    ).build()
def admin_menu():
    # پنل اصلی مدیر به دو بخش واضح تقسیم شده است:
    # ۱) امور مشتریان و سفارش‌ها
    # ۲) مدیریت محصولات
    # گزارش Excel نیز مستقیماً در همین پنل اصلی در دسترس است.
    return (ChatKeypadBuilder()
      .row(
          ChatKeypadBuilder().button(id="admin_customer_panel",text="👤 امور مشتریان"),
          ChatKeypadBuilder().button(id="admin_products",text="🛠 مدیریت محصولات")
      )
      .row(ChatKeypadBuilder().button(id="admin_export_orders",text="📊 دریافت گزارش Excel"))
      .build())
def admin_customer_menu():
    return (ChatKeypadBuilder()
      .row(ChatKeypadBuilder().button(id="admin_orders",text="📋 مشاهده سفارش‌ها"))
      .row(ChatKeypadBuilder().button(id="admin_after_sales",text="🔄 تعویض و بازگشت وجه"))
      .row(ChatKeypadBuilder().button(id="admin_export_orders",text="📊 دریافت گزارش Excel"))
      .row(ChatKeypadBuilder().button(id="admin_back",text="🔙 بازگشت به پنل اصلی"))
      .build())
def admin_products_menu():
    p=admin_product(); active=p[7] if p else 0
    t="🔴 غیرفعال کردن P-001" if active else "🟢 فعال کردن P-001"
    return (ChatKeypadBuilder().row(ChatKeypadBuilder().button(id="admin_price_P-001",text="💰 تغییر قیمت P-001"))
      .row(ChatKeypadBuilder().button(id="admin_name_P-001",text="📝 تغییر نام P-001"))
      .row(ChatKeypadBuilder().button(id="admin_description_P-001",text="📄 تغییر توضیحات P-001"))
      .row(ChatKeypadBuilder().button(id="admin_image_P-001",text="🖼 تغییر تصویر P-001"))
      .row(ChatKeypadBuilder().button(id="admin_delivery_P-001",text="⏱ تغییر زمان ارسال P-001"))
      .row(ChatKeypadBuilder().button(id="admin_toggle_P-001",text=t))
      .row(ChatKeypadBuilder().button(id="admin_product_info_P-001",text="📋 مشخصات کامل P-001"))
      .row(ChatKeypadBuilder().button(id="admin_back",text="🔙 بازگشت")).build())
def rating_menu(oid):
    return (ChatKeypadBuilder().row(*[ChatKeypadBuilder().button(id=f"rating_{oid}_{i}",text=f"⭐ {i}") for i in (1,2,3)])
      .row(ChatKeypadBuilder().button(id=f"rating_{oid}_4",text="⭐ 4"),ChatKeypadBuilder().button(id=f"rating_{oid}_5",text="⭐ 5")).build())


def build_product_text():
    p=product()
    if not p: return "❌ محصول موردنظر پیدا نشد."
    code,name,desc,pr,days,_,_=p
    cursor.execute("SELECT COUNT(*) FROM orders WHERE product=?",(code,)); total=cursor.fetchone()[0]
    since=(datetime.now()-timedelta(days=RECENT_ORDER_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("SELECT COUNT(*) FROM orders WHERE product=? AND created_at>=?",(code,since)); recent=cursor.fetchone()[0]
    cursor.execute("SELECT AVG(rating),COUNT(*) FROM reviews WHERE product=?",(code,)); avg,count=cursor.fetchone(); avg=round(avg or 0,1)
    text=(f"🏪 فروشگاه آنلاین شاپ\nخریدی آسان و مطمئن برای کلیه همشهریان و هموطنان\n\n📦 {code}\n🔹 نام محصول: {name}\n\n📝 توضیحات محصول:\n{desc}\n\n💰 قیمت: {price(pr)} تومان\n\n🛒 تاکنون {total} سفارش برای این محصول ثبت شده است.\n")
    if recent>=RECENT_ORDER_THRESHOLD: text+="🔥 این محصول اخیراً مورد توجه مشتریان بوده است.\n"
    return text+f"\n🚚 ارسال رایگان فقط در محدوده شهر {FREE_DELIVERY_CITY}\n⏱ زمان تقریبی ارسال: {days} روز کاری\n\n💵 پرداخت درب منزل پس از رویت و تحویل کالا\n\n⭐ امتیاز مشتریان: {avg} / 5\n💬 تعداد نظر ثبت شده: {count}\n\nبرای ثبت سفارش روی دکمه زیر بزنید."

async def show_product(chat_id):
    p=product()
    if not p: return await bot.send_message(chat_id=chat_id,text="❌ محصولی برای نمایش وجود ندارد.")
    text=build_product_text(); path=p[5]; fid=p[6]
    if path and os.path.exists(path):
        try:
            await bot.send_image(chat_id=chat_id,path=path,text=text,chat_keypad=product_menu(),chat_keypad_type="New"); return
        except Exception as e: print("IMAGE ERROR",e)
    if fid:
        try:
            await bot.send_image(chat_id=chat_id,file_id=fid,text=text,chat_keypad=product_menu(),chat_keypad_type="New"); return
        except Exception as e: print("FILE IMAGE ERROR",e)
    await bot.send_message(chat_id=chat_id,text="⚠️ تصویر محصول ثبت نشده است.\n\n"+text,chat_keypad=product_menu(),chat_keypad_type="New")

def policy_text():
    return ("📜 شرایط خدمات پس از فروش\n\n"
      f"1️⃣ مشتری هنگام تحویل کالا آن را رویت و بررسی می‌کند.\n\n"
      f"2️⃣ پس از تأیید دریافت، مهلت عادی تعویض/مرجوعی {RETURN_WINDOW_HOURS} ساعت است.\n\n"
      f"3️⃣ درخواست تعویض تا {EXCHANGE_WINDOW_HOURS} ساعت پس از تأیید دریافت قابل ثبت است.\n\n"
      "4️⃣ درخواست توسط مدیر بررسی و نتیجه به مشتری اعلام می‌شود.\n\n"
      "5️⃣ در صورت تأیید مرجوعی، اطلاعات بانکی مشتری دریافت و پس از بررسی مدیر، وضعیت واریز ثبت می‌شود.\n\n"
      "6️⃣ دکمه خدمات پس از فروش ممکن است پس از پایان مهلت نیز دیده شود، اما درخواست جدید پذیرفته نمی‌شود.")

async def notify_customer_status(oid,status):
    cursor.execute("SELECT chat_id,name,product FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
    if r and r[0]:
        await bot.send_message(chat_id=r[0],text=f"📢 وضعیت سفارش #{oid} تغییر کرد.\n\n📦 محصول: {r[2]}\n📌 وضعیت: {status_fa(status)}")
async def notify_admin_order(oid):
    cursor.execute("SELECT name,phone,address,product,status FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
    if not r:return
    await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"🆕 سفارش جدید #{oid}\n\n👤 نام: {r[0]}\n📱 تلفن: {r[1]}\n📍 آدرس: {r[2]}\n📦 محصول: {r[3]}\n📌 وضعیت: {status_fa(r[4])}",chat_keypad=order_admin_menu(oid),chat_keypad_type="New")
async def notify_after_sales(rid):
    cursor.execute("""SELECT r.order_id,r.request_type,r.reason,r.status,r.user_id,o.name,o.phone,o.address,o.product,r.refund_amount,r.exchange_product,r.card_number,r.iban,r.bank_name,r.exchange_delivery_days FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=?""",(rid,))
    x=cursor.fetchone()
    if not x:
        print("AFTER SALES NOTIFICATION ERROR: request not found", rid)
        return False
    oid,typ,reason,st,uid,name,phone,address,prod,amount,exprod,card,iban,bank_name,days=x
    text=(f"🔔 درخواست خدمات پس از فروش #{rid}\n\n"
          f"🧾 سفارش: #{oid}\n"
          f"👤 مشتری: {name}\n"
          f"📱 تلفن: {phone}\n"
          f"📍 آدرس: {address}\n"
          f"📦 محصول: {prod}\n"
          f"🔄 نوع درخواست: {typ}\n"
          f"📝 دلیل: {reason}\n"
          f"💰 مبلغ: {price(amount)} تومان\n"
          f"🔁 کالای جایگزین: {exprod or '-'}\n"
          f"💳 کارت: {card or '-'}\n"
          f"🏦 شبا: {iban or '-'}\n"
          f"🏦 بانک: {bank_name or '-'}\n"
          f"📌 وضعیت: {after_status(st)}")
    # Send the notification text separately from the management keypad. This prevents a
    # keypad validation error from blocking the actual request notification.
    try:
        await bot.send_message(chat_id=ADMIN_CHAT_ID,text=text)
        print("AFTER SALES ADMIN TEXT SENT", rid)
    except Exception as e:
        print("AFTER SALES ADMIN TEXT ERROR:", type(e).__name__, e)
        return False
    try:
        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=f"🎛 مدیریت درخواست #{rid} — {typ}\nوضعیت فعلی: {after_status(st)}\nلطفاً یکی از گزینه‌های زیر را انتخاب کنید:",
            chat_keypad=return_admin_menu(rid,typ,st),
            chat_keypad_type="New"
        )
        print("AFTER SALES ADMIN CONTROL SENT", rid)
    except Exception as e:
        print("AFTER SALES ADMIN CONTROL ERROR:", type(e).__name__, e)
        # Text has already reached the manager; report the control-keypad problem in CMD.
    return True

async def notify_customer_after(rid):
    cursor.execute("SELECT r.order_id,r.request_type,r.status,r.admin_note,r.user_id,o.chat_id,r.exchange_delivery_days,r.exchange_round FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=?",(rid,)); r=cursor.fetchone()
    if not r or not r[5]: return
    oid,typ,st,note,uid,cid,days,round_no=r
    text=f"📢 وضعیت درخواست #{rid} برای سفارش #{oid}\n\n🔄 نوع: {typ}\n📌 وضعیت: {after_status(st)}\n📝 توضیح: {note or '-'}"
    if typ=="تعویض" and st=="در حال تعویض":
        text += f"\n\n📦 درخواست تعویض شما تأیید شد.\n⏱ زمان تقریبی ارسال کالای تعویضی: {days or '-'} روز کاری\n\nپس از ارسال کالا، دکمه تأیید دریافت برای شما فعال می‌شود."
    elif typ=="تعویض" and st=="کالای تعویضی ارسال شد":
        text += "\n\n🚚 کالای تعویضی ارسال شد.\nلطفاً پس از دریافت و بررسی، دریافت کالای تعویضی را تأیید کنید."
    elif typ=="تعویض" and st=="تعویض انجام شد":
        text += "\n\n✅ تعویض این سفارش با موفقیت تکمیل شد.\nلطفاً از منوی بعدی رضایت خود را ثبت کنید یا در صورت نارضایتی درخواست تعویض مجدد/مرجوعی بدهید."
    elif typ=="مرجوعی و بازگشت وجه" and st=="در انتظار واریز":
        text += "\n\n✅ درخواست مرجوعی شما توسط مدیر تأیید شد.\n💰 مبلغ کالا حداکثر تا ۱ روز کاری به حساب اعلام‌شده واریز خواهد شد.\n\nپس از انجام واریز، پیام دیگری برای تأیید دریافت وجه برای شما ارسال می‌شود."
    elif typ=="مرجوعی و بازگشت وجه" and st=="وجه واریز شد":
        text += "\n\n💰 وجه کالا واریز شده است.\nلطفاً پس از بررسی حساب خود، دریافت وجه را تأیید کنید."
    elif typ=="مرجوعی و بازگشت وجه" and st=="تکمیل شده":
        text += "\n\n🙏 دریافت وجه شما با موفقیت تأیید شد. درخواست مرجوعی تکمیل گردید."
    if typ=="تعویض" and st=="کالای تعویضی ارسال شد":
        await bot.send_message(chat_id=cid,text=text,chat_keypad=exchange_accept_menu(rid),chat_keypad_type="New")
    elif typ=="مرجوعی و بازگشت وجه" and st=="وجه واریز شد":
        await bot.send_message(chat_id=cid,text=text,chat_keypad=refund_accept_menu(rid),chat_keypad_type="New")
    else:
        await bot.send_message(chat_id=cid,text=text)

def export_excel(chat_id):
    if not OPENPYXL_AVAILABLE:
        return bot.send_message(chat_id=chat_id,text="❌ openpyxl نصب نیست. دستور: pip install openpyxl")
    path="orders_export.xlsx"; wb=Workbook(); ws=wb.active; ws.title="Orders"
    headers=["شماره سفارش","کاربر","محصول","نام","تلفن","آدرس","وضعیت","ثبت","تحویل","تأیید دریافت","امتیاز","نظر"]
    ws.append(headers)
    cursor.execute("""SELECT o.id,o.user_id,o.product,o.name,o.phone,o.address,o.status,o.created_at,o.delivered_at,o.accepted_at,rv.rating,rv.comment FROM orders o LEFT JOIN reviews rv ON rv.order_id=o.id ORDER BY o.id DESC""")
    for row in cursor.fetchall(): ws.append(list(row))
    ws2=wb.create_sheet("AfterSales"); ws2.append(["شناسه درخواست","سفارش","کاربر","نوع","دلیل","وضعیت","مبلغ","کارت","شبا","بانک","کالای جایگزین","ثبت","تصمیم","واریز"])
    cursor.execute("SELECT id,order_id,user_id,request_type,reason,status,refund_amount,card_number,iban,bank_name,exchange_product,created_at,decided_at,paid_at FROM return_requests ORDER BY id DESC")
    for row in cursor.fetchall(): ws2.append(list(row))
    for sh in wb.worksheets:
        for c in sh[1]: c.font=Font(bold=True); c.alignment=Alignment(horizontal="center")
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width=min(max(max(len(str(x.value or "")) for x in col)+2,10),35)
    wb.save(path)
    return bot.send_document(chat_id=chat_id,path=path,caption="📊 گزارش کامل سفارش‌ها و خدمات پس از فروش")

async def show_my_orders(message):
    uid=get_user_id(message); cid=get_chat_id(message)
    cursor.execute("SELECT id,product,status,created_at,accepted_at FROM orders WHERE user_id=? ORDER BY id DESC",(uid,)); rows=cursor.fetchall()
    if not rows:return await bot.send_message(chat_id=cid,text="📋 هنوز سفارشی برای شما ثبت نشده است.",chat_keypad=main_menu(),chat_keypad_type="New")
    for oid,prod,st,created,accepted in rows:
        txt=f"🧾 سفارش #{oid}\n📦 محصول: {prod}\n📌 وضعیت: {status_fa(st)}\n📅 ثبت: {created or '-'}"
        keypad=None
        if st=="تحویل و تأیید شده":
            d=deadline(accepted)
            exchange_deadline=deadline(accepted,EXCHANGE_WINDOW_HOURS)
            if exchange_deadline and datetime.now()<=exchange_deadline: keypad=customer_after_menu(oid)
            else: txt+=f"\n⏰ مهلت تعویض تا: {exchange_deadline.strftime('%Y-%m-%d %H:%M:%S') if exchange_deadline else '-'}"
            cursor.execute("SELECT id,status,request_type FROM return_requests WHERE order_id=?",(oid,)); rr=cursor.fetchone()
            if rr: txt+=f"\n\n🔄 درخواست خدمات پس از فروش: #{rr[0]}\n📌 {after_status(rr[1])}\nنوع: {rr[2]}"
        cursor.execute("SELECT id,status,request_type,exchange_delivery_days FROM return_requests WHERE order_id=?",(oid,)); rr=cursor.fetchone()
        if rr and rr[1]=="کالای تعویضی ارسال شد":
            txt += f"\n\n🚚 کالای تعویضی ارسال شده است.\n⏱ زمان تقریبی تحویل: {rr[3] or '-'} روز کاری"
            keypad = exchange_accept_menu(rr[0])
        elif rr and rr[1]=="تعویض انجام شد":
            txt += "\n\n✅ تعویض انجام شده است. از منوی رضایت/خدمات پس از فروش اقدام کنید."
            keypad = customer_after_menu(oid, exchange=True)
        if keypad:
            await bot.send_message(chat_id=cid,text=txt,chat_keypad=keypad,chat_keypad_type="New")
        else:
            await bot.send_message(chat_id=cid,text=txt)

@bot.on_message()
async def handle_message(bot_instance: Robot,message: Message):
    uid=get_user_id(message); cid=get_chat_id(message); text=(getattr(message,"text","") or "").strip()
    if text=="/start":
        if cid==ADMIN_CHAT_ID:
            return await bot.send_message(chat_id=cid,text="👨‍💼 پنل مدیریت فروشگاه",chat_keypad=admin_menu(),chat_keypad_type="New")
        return await bot.send_message(chat_id=cid,text="🏪 فروشگاه آنلاین شاپ\n\nلطفاً گزینه موردنظر را انتخاب کنید.",chat_keypad=main_menu(),chat_keypad_type="New")
    if cid==ADMIN_CHAT_ID and cid in admin_states:
        state=admin_states[cid]
        if state=="waiting_image":
            raw=getattr(message,"raw_data",None) or {}
            def find_fid(x):
                if isinstance(x,dict):
                    for k in ("file_id","fileId"):
                        if x.get(k): return str(x[k])
                    for v in x.values():
                        z=find_fid(v)
                        if z:return z
                elif isinstance(x,list):
                    for v in x:
                        z=find_fid(v)
                        if z:return z
                return None
            fid=find_fid(raw)
            if not fid:
                return await bot.send_message(chat_id=cid,text="❌ شناسه تصویر دریافت نشد. لطفاً عکس را دوباره ارسال کنید.")
            tmp=PRODUCT_IMAGE_PATH+".new"
            try:
                await bot.download(file_id=fid,save_as=tmp,verbose=False)
                if not os.path.exists(tmp) or os.path.getsize(tmp)<=0: raise RuntimeError("empty image")
                if os.path.exists(PRODUCT_IMAGE_PATH): os.remove(PRODUCT_IMAGE_PATH)
                os.replace(tmp,PRODUCT_IMAGE_PATH)
                cursor.execute("UPDATE products SET image_path=?,image_file_id=NULL WHERE code=?",(PRODUCT_IMAGE_PATH,PRODUCT_CODE)); conn.commit()
                admin_states.pop(cid,None)
                return await bot.send_message(chat_id=cid,text="✅ تصویر محصول با موفقیت تغییر کرد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            except Exception as e:
                print("IMAGE UPDATE ERROR:",repr(e))
                try:
                    if os.path.exists(tmp): os.remove(tmp)
                except: pass
                return await bot.send_message(chat_id=cid,text="❌ ذخیره تصویر انجام نشد. لطفاً دوباره تلاش کنید.")
        if state.startswith("waiting_"):
            if text=="لغو": admin_states.pop(cid,None); return await bot.send_message(chat_id=cid,text="❌ عملیات لغو شد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            field=state.replace("waiting_","")
            if field=="price" and not norm(text).isdigit(): return await bot.send_message(chat_id=cid,text="❌ قیمت باید فقط عدد باشد.")
            dbfield={"price":"price","name":"name","description":"description","delivery":"delivery_days"}.get(field)
            if dbfield:
                cursor.execute(f"UPDATE products SET {dbfield}=? WHERE code=?",(norm(text) if field=="price" else text,PRODUCT_CODE)); conn.commit(); admin_states.pop(cid,None)
                return await bot.send_message(chat_id=cid,text="✅ اطلاعات محصول تغییر کرد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
    if uid in customers:
        st=customers[uid]
        if st["step"]=="name": st["name"]=text; st["step"]="phone"; return await bot.send_message(chat_id=cid,text="📱 شماره تماس خود را وارد کنید:")
        if st["step"]=="phone": st["phone"]=norm(text); st["step"]="address"; return await bot.send_message(chat_id=cid,text="📍 آدرس کامل را وارد کنید:")
        if st["step"]=="address":
            st["address"]=text; st["step"]="confirm"
            return await bot.send_message(chat_id=cid,text=f"🧾 اطلاعات سفارش\n\n👤 {st['name']}\n📱 {st['phone']}\n📍 {st['address']}\n📦 {st['product']}\n\nآیا سفارش تأیید شود؟",chat_keypad=confirm_menu(),chat_keypad_type="New")
    if uid in after_sales_states:
        st=after_sales_states[uid]
        if text=="لغو": after_sales_states.pop(uid,None); return await bot.send_message(chat_id=cid,text="❌ عملیات لغو شد.",chat_keypad=main_menu(),chat_keypad_type="New")
        if st["step"]=="reason_again":
            st["reason"]=text; rid=st["rid"]; typ=st["type"]
            if typ=="تعویض":
                cursor.execute("UPDATE return_requests SET status='در انتظار بررسی مجدد',reason=?,decided_at=NULL,admin_note=NULL,exchange_round=COALESCE(exchange_round,1)+1 WHERE id=?",(text,rid)); conn.commit(); after_sales_states.pop(uid,None)
                await bot.send_message(chat_id=cid,text=f"✅ درخواست تعویض مجدد #{rid} ثبت شد و برای مدیر ارسال شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_after_sales(rid)
            st["step"]="bank"
            return await bot.send_message(chat_id=cid,text="🏦 اطلاعات بازگشت وجه\n\nنام صاحب حساب را وارد کنید:")
        if st["step"]=="reason":
            st["reason"]=text; oid=st["order_id"]
            cursor.execute("SELECT user_id,product,name,phone FROM orders WHERE id=?",(oid,)); o=cursor.fetchone()
            if not o or o[0]!=uid: after_sales_states.pop(uid,None); return
            amount=None; cursor.execute("SELECT price FROM products WHERE code=?",(o[1],)); pr=cursor.fetchone(); amount=pr[0] if pr else ""
            if st["type"]=="تعویض":
                cursor.execute("INSERT OR IGNORE INTO return_requests(order_id,user_id,product,request_type,reason,status,created_at,return_deadline,customer_name,phone,refund_amount) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(oid,uid,o[1],"تعویض",text,"در انتظار بررسی",now_string(),deadline(st.get("accepted_at") or now_string(),EXCHANGE_WINDOW_HOURS).strftime("%Y-%m-%d %H:%M:%S"),o[2],o[3],amount)); conn.commit()
                rid=cursor.lastrowid; after_sales_states.pop(uid,None); await bot.send_message(chat_id=cid,text=f"✅ درخواست تعویض #{rid} ثبت شد و برای مدیر ارسال شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_after_sales(rid)
            st["step"]="bank"; return await bot.send_message(chat_id=cid,text="🏦 اطلاعات بازگشت وجه\n\nنام صاحب حساب را وارد کنید:")
        if st["step"]=="bank": st["account_name"]=text; st["step"]="card"; return await bot.send_message(chat_id=cid,text="💳 شماره کارت ۱۶ رقمی را وارد کنید:")
        if st["step"]=="card":
            c=norm(text).replace("-","").replace(" ","")
            if len(c)!=16 or not c.isdigit(): return await bot.send_message(chat_id=cid,text="❌ شماره کارت باید ۱۶ رقم باشد.")
            st["card"]=c; st["step"]="iban"; return await bot.send_message(chat_id=cid,text="🏦 شماره شبا را وارد کنید (مثلاً IR...):")
        if st["step"]=="iban":
            iban=norm(text).replace(" ","").upper()
            if iban.startswith("IR"): iban=iban
            st["iban"]=iban; st["step"]="bank_name"; return await bot.send_message(chat_id=cid,text="🏦 نام بانک را وارد کنید:")
        if st["step"]=="bank_name":
            st["bank_name"]=text; oid=st["order_id"]
            cursor.execute("SELECT user_id,product,name,phone FROM orders WHERE id=?",(oid,)); o=cursor.fetchone()
            cursor.execute("SELECT price FROM products WHERE code=?",(o[1],)); pr=cursor.fetchone(); amount=pr[0] if pr else ""
            # Each order has one active after-sales record in the current database schema.
            # If a previous exchange record exists, reuse it for the refund request instead of
            # silently ignoring the INSERT because of the UNIQUE(order_id) constraint.
            cursor.execute("SELECT id FROM return_requests WHERE order_id=?",(oid,)); existing=cursor.fetchone()
            if existing:
                rid=existing[0]
                cursor.execute("""UPDATE return_requests SET user_id=?,product=?,request_type=?,reason=?,status=?,created_at=?,return_deadline=?,customer_name=?,phone=?,refund_amount=?,bank_account_name=?,card_number=?,iban=?,bank_name=?,decided_at=NULL,admin_note=NULL,paid_at=NULL WHERE id=?""",(uid,o[1],"مرجوعی و بازگشت وجه",st["reason"],"در انتظار بررسی",now_string(),deadline(st.get("accepted_at") or now_string(),RETURN_WINDOW_HOURS).strftime("%Y-%m-%d %H:%M:%S"),o[2],o[3],amount,st["account_name"],st["card"],st["iban"],text,rid))
            else:
                cursor.execute("INSERT INTO return_requests(order_id,user_id,product,request_type,reason,status,created_at,return_deadline,customer_name,phone,refund_amount,bank_account_name,card_number,iban,bank_name) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(oid,uid,o[1],"مرجوعی و بازگشت وجه",st["reason"],"در انتظار بررسی",now_string(),deadline(st.get("accepted_at") or now_string(),RETURN_WINDOW_HOURS).strftime("%Y-%m-%d %H:%M:%S"),o[2],o[3],amount,st["account_name"],st["card"],st["iban"],text))
                rid=cursor.lastrowid
            conn.commit()
            after_sales_states.pop(uid,None)
            sent=await bot.send_message(chat_id=cid,text=f"✅ درخواست بازگشت وجه #{rid} ثبت شد و برای مدیر ارسال شد.",chat_keypad=main_menu(),chat_keypad_type="New")
            ok=await notify_after_sales(rid)
            if not ok:
                await bot.send_message(chat_id=cid,text="⚠️ درخواست شما در سامانه ثبت شد، اما ارسال اعلان به مدیر با خطا مواجه شد. مدیر می‌تواند درخواست را از پنل خدمات پس از فروش مشاهده کند.")
            return
    if rating_comments.get(uid):
        data=rating_comments.pop(uid); cursor.execute("INSERT INTO reviews(order_id,user_id,product,rating,comment,created_at) VALUES(?,?,?,?,?,?)",(data["order_id"],uid,data["product"],data["rating"],"" if text=="بدون نظر" else text,now_string())); conn.commit(); return await bot.send_message(chat_id=cid,text="🙏 ممنون از ثبت نظر شما.",chat_keypad=main_menu(),chat_keypad_type="New")

@bot.on_callback()
async def handle_callback(bot_instance: Robot,message: Message):
    bid=button_id(message); uid=get_user_id(message); cid=get_chat_id(message)
    if not bid:return
    if bid.startswith("admin_"):
        if not is_admin(message):return
        if bid=="admin_back": return await bot.send_message(chat_id=cid,text="👨‍💼 پنل مدیریت فروشگاه",chat_keypad=admin_menu(),chat_keypad_type="New")
        if bid=="admin_customer_panel":
            return await bot.send_message(
                chat_id=cid,
                text="👤 بخش امور مشتریان و سفارش‌ها\n\nاز این قسمت سفارش‌ها، خدمات پس از فروش و گزارش Excel را مدیریت کنید.",
                chat_keypad=admin_customer_menu(),
                chat_keypad_type="New"
            )
        if bid=="admin_products": return await bot.send_message(chat_id=cid,text="🛠 مدیریت محصولات",chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid=="admin_after_sales":
            cursor.execute("SELECT id,order_id,request_type,status FROM return_requests WHERE status NOT IN ('تکمیل شده','تعویض انجام شد','وجه واریز شد','رد شده') ORDER BY id DESC")
            rows=cursor.fetchall()
            if not rows:return await bot.send_message(chat_id=cid,text="✅ درخواست باز خدمات پس از فروش وجود ندارد.",chat_keypad=admin_menu(),chat_keypad_type="New")
            for rid,oid,typ,st in rows: await bot.send_message(chat_id=cid,text=f"🔄 درخواست #{rid}\n🧾 سفارش #{oid}\nنوع: {typ}\nوضعیت: {after_status(st)}",chat_keypad=return_admin_menu(rid,typ,st),chat_keypad_type="New")
            return
        if bid=="admin_export_orders": return await export_excel(cid)
        if bid=="admin_orders":
            cursor.execute("SELECT id,name,product,status FROM orders ORDER BY id DESC LIMIT 20")
            rows=cursor.fetchall(); txt="📋 آخرین سفارش‌ها\n\n"+"\n".join(f"#{r[0]} | {r[1]} | {r[2]} | {status_fa(r[3])}" for r in rows) if rows else "📋 سفارشی وجود ندارد."
            return await bot.send_message(chat_id=cid,text=txt,chat_keypad=admin_menu(),chat_keypad_type="New")
        if bid in ("admin_price_P-001","admin_name_P-001","admin_description_P-001","admin_delivery_P-001"):
            admin_states[cid]="waiting_"+bid.replace("admin_","").replace("_P-001","")
            return await bot.send_message(chat_id=cid,text="✏️ مقدار جدید را در پیام بعدی ارسال کنید. برای لغو: لغو")
        if bid=="admin_image_P-001":
            admin_states[cid]="waiting_image"
            return await bot.send_message(chat_id=cid,text="🖼 تصویر جدید محصول را به صورت عکس ارسال کنید.\nبرای لغو: لغو")
        if bid=="admin_toggle_P-001":
            p=admin_product(); new=0 if p[7] else 1; cursor.execute("UPDATE products SET active=? WHERE code=?",(new,PRODUCT_CODE)); conn.commit(); return await bot.send_message(chat_id=cid,text="✅ وضعیت محصول تغییر کرد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid=="admin_product_info_P-001":
            p=admin_product(); return await bot.send_message(chat_id=cid,text=f"📋 {p[0]}\nنام: {p[1]}\nتوضیحات: {p[2]}\nقیمت: {price(p[3])} تومان\nارسال: {p[4]} روز کاری\nوضعیت: {'فعال' if p[7] else 'غیرفعال'}",chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid.startswith("admin_confirm_"):
            oid=int(bid.split("_")[-1]); cursor.execute("UPDATE orders SET status='تأیید شده' WHERE id=?",(oid,)); conn.commit(); await bot.send_message(chat_id=cid,text=f"✅ سفارش #{oid} تأیید شد."); return await notify_customer_status(oid,"تأیید شده")
        if bid.startswith("admin_cancel_"):
            oid=int(bid.split("_")[-1]); cursor.execute("UPDATE orders SET status='لغو شده' WHERE id=?",(oid,)); conn.commit(); await bot.send_message(chat_id=cid,text=f"❌ سفارش #{oid} لغو شد."); return await notify_customer_status(oid,"لغو شده")
        if bid.startswith("admin_deliver_"):
            oid=int(bid.split("_")[-1]); cursor.execute("SELECT chat_id,name,product FROM orders WHERE id=?",(oid,)); r=cursor.fetchone();
            if not r:return
            cursor.execute("UPDATE orders SET status='تحویل شده',delivered_at=? WHERE id=?",(now_string(),oid)); conn.commit(); await bot.send_message(chat_id=cid,text=f"🚚 سفارش #{oid} تحویل شد."); return await bot.send_message(chat_id=r[0],text=f"سلام {r[1]} عزیز 🌷\n\n📦 سفارش #{oid} تحویل شما شده است. پس از رویت و بررسی، دریافت کالا را تأیید کنید.",chat_keypad=accept_menu(oid),chat_keypad_type="New")
        if bid.startswith("admin_refund_paid_"):
            rid=int(bid.split("_")[-1])
            cursor.execute("SELECT status FROM return_requests WHERE id=? AND request_type='مرجوعی و بازگشت وجه'",(rid,)); r=cursor.fetchone()
            if not r or r[0]!="در انتظار واریز": return
            cursor.execute("UPDATE return_requests SET status='وجه واریز شد',paid_at=?,decided_at=? WHERE id=?",(now_string(),now_string(),rid)); conn.commit()
            await bot.send_message(chat_id=cid,text=f"✅ واریز درخواست #{rid} ثبت شد.")
            return await notify_customer_after(rid)
        if bid.startswith("admin_exchange_sent_"):
            rid=int(bid.split("_")[-1])
            cursor.execute("SELECT status FROM return_requests WHERE id=? AND request_type='تعویض'",(rid,)); r=cursor.fetchone()
            if not r or r[0]!="در حال تعویض": return
            cursor.execute("UPDATE return_requests SET status='کالای تعویضی ارسال شد',exchange_sent_at=?,decided_at=? WHERE id=?",(now_string(),now_string(),rid)); conn.commit()
            await bot.send_message(chat_id=cid,text=f"📦 ارسال کالای تعویضی برای درخواست #{rid} ثبت شد.")
            return await notify_customer_after(rid)
        if bid.startswith("admin_exchange_done_"):
            rid=int(bid.split("_")[-1])
            cursor.execute("SELECT status FROM return_requests WHERE id=? AND request_type='تعویض'",(rid,)); r=cursor.fetchone()
            if not r or r[0]!="در انتظار تأیید دریافت تعویضی": return
            cursor.execute("UPDATE return_requests SET status='تعویض انجام شد',decided_at=? WHERE id=?",(now_string(),rid)); conn.commit()
            await bot.send_message(chat_id=cid,text=f"✅ تعویض نهایی درخواست #{rid} ثبت شد.")
            return await notify_customer_after(rid)
        if bid.startswith("admin_after_approve_") or bid.startswith("admin_after_reject_"):
            rid=int(bid.split("_")[-1]); approve="approve" in bid
            cursor.execute("SELECT order_id,status,request_type FROM return_requests WHERE id=?",(rid,)); r=cursor.fetchone()
            if not r:return
            new="رد شده" if not approve else ("در انتظار واریز" if r[2]=="مرجوعی و بازگشت وجه" else "در حال تعویض")
            if approve and r[2]=="تعویض":
                cursor.execute("UPDATE return_requests SET status=?,decided_at=?,admin_note=?,exchange_delivery_days=? WHERE id=?",(new,now_string(),"درخواست تعویض توسط مدیر تأیید شد؛ منتظر ارسال کالای تعویضی.","2 تا 3",rid))
            else:
                cursor.execute("UPDATE return_requests SET status=?,decided_at=?,admin_note=? WHERE id=?",(new,now_string(),"درخواست توسط مدیر تأیید شد." if approve else "درخواست توسط مدیر رد شد.",rid))
            conn.commit()
            await bot.send_message(chat_id=cid,text=f"{'✅' if approve else '❌'} درخواست #{rid} بررسی شد.")
            # After approval, always send a fresh admin control message for the next step.
            if approve and r[2] == "تعویض":
                await bot.send_message(
                    chat_id=cid,
                    text=(f"📦 درخواست تعویض #{rid} تأیید شد.\n\n"
                          f"⏱ زمان تقریبی ارسال کالای تعویضی: 2 تا 3 روز کاری\n\n"
                          "پس از آماده شدن و ارسال کالا، دکمه زیر را بزنید:"),
                    chat_keypad=return_admin_menu(rid, "تعویض", "در حال تعویض"),
                    chat_keypad_type="New"
                )
            elif approve and r[2] == "مرجوعی و بازگشت وجه":
                await bot.send_message(
                    chat_id=cid,
                    text=(f"💰 درخواست مرجوعی #{rid} تأیید شد.\n\n"
                          "⏱ مبلغ کالا حداکثر تا ۱ روز کاری به حساب مشتری واریز خواهد شد.\n\n"
                          "پس از انجام واقعی واریز، دکمه زیر را بزنید تا پیام تأیید دریافت وجه برای مشتری ارسال شود:"),
                    chat_keypad=return_admin_menu(rid, "مرجوعی و بازگشت وجه", "در انتظار واریز"),
                    chat_keypad_type="New"
                )
            return await notify_customer_after(rid)
        return
    if bid=="main_menu":
        customers.pop(uid,None); after_sales_states.pop(uid,None); return await bot.send_message(chat_id=cid,text="🏪 فروشگاه آنلاین شاپ\n\nلطفاً گزینه موردنظر را انتخاب کنید.",chat_keypad=main_menu(),chat_keypad_type="New")
    if bid=="show_products": return await show_product(cid)
    if bid=="shop_policy": return await bot.send_message(chat_id=cid,text=policy_text(),chat_keypad=main_menu(),chat_keypad_type="New")
    if bid=="my_orders": return await show_my_orders(message)
    if bid=="order_product_P-001":
        if not product(): return await bot.send_message(chat_id=cid,text="❌ محصول موجود نیست.")
        customers[uid]={"step":"name","product":PRODUCT_CODE,"name":"","phone":"","address":""}; return await bot.send_message(chat_id=cid,text="🛒 ثبت سفارش\n\nنام و نام خانوادگی خود را وارد کنید:")
    if bid=="customer_confirm_order":
        st=customers.get(uid)
        if not st or st.get("step")!="confirm":return
        cursor.execute("INSERT INTO orders(user_id,product,name,phone,address,status,chat_id,created_at) VALUES(?,?,?,?,?,?,?,?)",(uid,st["product"],st["name"],st["phone"],st["address"],"جدید",cid,now_string())); conn.commit(); oid=cursor.lastrowid; customers.pop(uid,None); await bot.send_message(chat_id=cid,text=f"✅ سفارش #{oid} با موفقیت ثبت شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_admin_order(oid)
    if bid=="customer_cancel_order": customers.pop(uid,None); return await bot.send_message(chat_id=cid,text="❌ سفارش لغو شد.",chat_keypad=main_menu(),chat_keypad_type="New")
    if bid.startswith("customer_accept_delivery_"):
        oid=int(bid.split("_")[-1]); cursor.execute("SELECT user_id,status FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
        if not r or r[0]!=uid or r[1]!="تحویل شده":return
        accepted=now_string(); cursor.execute("UPDATE orders SET status='تحویل و تأیید شده',accepted_at=?,customer_acceptance=? WHERE id=?",(accepted,"مشتری کالا را رویت و سالم و مطابق سفارش تحویل گرفته است.",oid)); conn.commit(); d=deadline(accepted); await bot.send_message(chat_id=cid,text=f"✅ دریافت سفارش #{oid} ثبت شد.\n\n⏰ مهلت خدمات پس از فروش عادی تا: {d.strftime('%Y-%m-%d %H:%M:%S')}",chat_keypad=customer_after_menu(oid),chat_keypad_type="New"); await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"📦 مشتری دریافت سفارش #{oid} را تأیید کرد."); return
    if bid.startswith("customer_accept_refund_"):
        rid=int(bid.split("_")[-1])
        cursor.execute("SELECT r.order_id,r.user_id,r.status,o.chat_id FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=? AND r.request_type='مرجوعی و بازگشت وجه'",(rid,)); r=cursor.fetchone()
        if not r or r[1]!=uid or r[2]!="وجه واریز شد": return
        cursor.execute("UPDATE return_requests SET status='تکمیل شده',decided_at=?,admin_note=? WHERE id=?",(now_string(),"مشتری دریافت وجه را تأیید کرد.",rid)); conn.commit()
        oid=r[0]
        cursor.execute("UPDATE orders SET status='مرجوع و وجه بازگشت داده شد' WHERE id=?",(oid,)); conn.commit()
        await bot.send_message(chat_id=cid,text=f"✅ دریافت وجه درخواست #{rid} ثبت شد.\n\n🙏 ممنون از تأیید شما؛ فرآیند مرجوعی و بازگشت وجه تکمیل شد.",chat_keypad=main_menu(),chat_keypad_type="New")
        await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"💰 مشتری دریافت وجه درخواست #{rid} برای سفارش #{oid} را تأیید کرد.")
        return
    if bid.startswith("customer_accept_exchange_"):
        rid=int(bid.split("_")[-1])
        cursor.execute("SELECT r.order_id,r.user_id,r.status,o.chat_id FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=? AND r.request_type='تعویض'",(rid,)); r=cursor.fetchone()
        if not r or r[1]!=uid or r[2]!="کالای تعویضی ارسال شد": return
        cursor.execute("UPDATE return_requests SET status='تعویض انجام شد',exchange_received_at=?,exchange_customer_acceptance=?,decided_at=? WHERE id=?",(now_string(),"مشتری کالای تعویضی را دریافت و تأیید کرد.",now_string(),rid)); conn.commit()
        oid=r[0]
        cursor.execute("UPDATE orders SET status='تعویض انجام شد' WHERE id=?",(oid,)); conn.commit()
        await bot.send_message(chat_id=cid,text=f"✅ دریافت کالای تعویضی درخواست #{rid} ثبت شد.\n\nاکنون رضایت خود را اعلام کنید.",chat_keypad=customer_after_menu(oid, exchange=True),chat_keypad_type="New")
        await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"📦 مشتری دریافت کالای تعویضی درخواست #{rid} را تأیید کرد.",chat_keypad=return_admin_menu(rid,"تعویض","در انتظار تأیید دریافت تعویضی"),chat_keypad_type="New")
        return
    if bid.startswith("customer_exchange_") or bid.startswith("customer_refund_"):
        oid=int(bid.split("_")[-1]); typ="تعویض" if bid.startswith("customer_exchange_") else "مرجوعی و بازگشت وجه"
        cursor.execute("SELECT user_id,status,accepted_at FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
        if not r or r[0]!=uid or r[1] not in ("تحویل و تأیید شده","تعویض انجام شد"):return
        cursor.execute("SELECT id,status,request_type FROM return_requests WHERE order_id=?",(oid,)); ex=cursor.fetchone()
        if r[1]=="تحویل و تأیید شده":
            d=deadline(r[2],EXCHANGE_WINDOW_HOURS if typ=="تعویض" else RETURN_WINDOW_HOURS)
            if not d or datetime.now()>d:return await bot.send_message(chat_id=cid,text="⏰ مهلت ثبت این درخواست به پایان رسیده است.")
        if r[1]=="تعویض انجام شد" and not ex: return
        if ex and ex[1] not in ("تعویض انجام شد", "رد شده") and r[1]=="تحویل و تأیید شده":
            return await bot.send_message(chat_id=cid,text=f"⚠️ برای این سفارش قبلاً درخواست #{ex[0]} ثبت شده است.\nوضعیت: {after_status(ex[1])}")
        if ex and r[1]=="تعویض انجام شد":
            after_sales_states[uid]={"step":"reason_again","order_id":oid,"type":typ,"accepted_at":r[2],"rid":ex[0]}
            return await bot.send_message(chat_id=cid,text=f"🔄 درخواست {typ} مجدد\n\nدلیل درخواست را کامل بنویسید:\n\nبرای لغو: لغو")
        after_sales_states[uid]={"step":"reason","order_id":oid,"type":typ,"accepted_at":r[2]}; return await bot.send_message(chat_id=cid,text=f"🔄 درخواست {typ}\n\nدلیل درخواست را کامل بنویسید:\n\nبرای لغو: لغو")
    if bid.startswith("customer_rate_"):
        oid=int(bid.split("_")[-1])
        cursor.execute("SELECT user_id,status FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
        if not r or r[0]!=uid or r[1] not in ("تحویل و تأیید شده","تعویض انجام شد"): return
        cursor.execute("SELECT id FROM reviews WHERE order_id=?",(oid,))
        if cursor.fetchone():
            return await bot.send_message(chat_id=cid,text="⭐ امتیاز و نظر این سفارش قبلاً ثبت شده است.",chat_keypad=main_menu(),chat_keypad_type="New")
        return await bot.send_message(chat_id=cid,text=f"⭐ لطفاً میزان رضایت خود از سفارش #{oid} را انتخاب کنید:",chat_keypad=rating_menu(oid),chat_keypad_type="New")
    if bid.startswith("rating_"):
        parts=bid.split("_"); oid=int(parts[1]); rating=int(parts[2]); cursor.execute("SELECT user_id,product,status FROM orders WHERE id=?",(oid,)); r=cursor.fetchone()
        if not r or r[0]!=uid or r[2] not in ("تحویل و تأیید شده","تعویض انجام شد"):return
        cursor.execute("SELECT id FROM reviews WHERE order_id=?",(oid,));
        if cursor.fetchone():return
        rating_comments[uid]={"order_id":oid,"product":r[1],"rating":rating}; return await bot.send_message(chat_id=cid,text=f"⭐ امتیاز {rating} از 5 ثبت شد.\n\n📝 اگر مایل هستید نظر خود را درباره کالا بنویسید.\nاگر نظری ندارید، بنویسید: بدون نظر")

print("BOT STARTING...")
bot.run()
