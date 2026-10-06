from rubka import Robot, Message
from rubka.keypad import ChatKeypadBuilder
import sqlite3
import os
import asyncio
import threading
import json
import urllib.request
import base64
from datetime import datetime, timedelta

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ============================================================
# SECURITY / CONFIGURATION
# ============================================================
TOKEN = ""
ADMIN_CHAT_ID = ""

DB_PATH = "shop.db"
PRODUCT_IMAGE_PATH = "product_default.jpg"
FREE_DELIVERY_CITY = "تبریز"
RECENT_ORDER_DAYS = 7
RECENT_ORDER_THRESHOLD = 3
RETURN_WINDOW_HOURS = 24
EXCHANGE_WINDOW_HOURS = 72
WEBSITE_BASE_URL = "https://nateghihamed1984-droid.github.io/SABALAN-SHOP/"
WEBSITE_API_URL = "http://127.0.0.1:8000/api/orders"
WEBSITE_SYNC_INTERVAL = 3

bot = Robot(token=TOKEN)
BOT_LOOP = None

customers = {}
rating_comments = {}
admin_states = {}
after_sales_states = {}

# ============================================================
# DATABASE
# ============================================================
def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_dt(value):
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S") if value else None
    except Exception:
        return None


def norm(value):
    return str(value or "").translate(
        str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    )


def price(value):
    try:
        return f"{int(norm(value)):,}"
    except Exception:
        return str(value or "نامشخص")


def ensure_column(conn, table, name, definition):
    columns = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if name not in columns:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")


def init_db():
    conn = get_db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS categories(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_id INTEGER,
            name TEXT NOT NULL,
            active INTEGER DEFAULT 1,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS products(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT,
            description TEXT,
            price TEXT,
            delivery_days TEXT,
            image_path TEXT,
            image_file_id TEXT,
            category_id INTEGER,
            active INTEGER DEFAULT 1,
            created_at TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS orders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            chat_id TEXT,
            product TEXT,
            name TEXT,
            phone TEXT,
            province TEXT,
            city TEXT,
            address TEXT,
            status TEXT,
            created_at TEXT,
            delivered_at TEXT,
            accepted_at TEXT,
            customer_acceptance TEXT,
            website_order_id INTEGER,
            website_customer_ref TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS reviews(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER,
            user_id TEXT,
            product TEXT,
            rating INTEGER,
            comment TEXT,
            created_at TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS return_requests(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER UNIQUE,
            user_id TEXT,
            product TEXT,
            request_type TEXT,
            reason TEXT,
            status TEXT,
            created_at TEXT,
            decided_at TEXT,
            admin_note TEXT,
            return_deadline TEXT,
            customer_name TEXT,
            phone TEXT,
            exchange_product TEXT,
            exchange_difference TEXT,
            refund_amount TEXT,
            bank_account_name TEXT,
            card_number TEXT,
            iban TEXT,
            bank_name TEXT,
            paid_at TEXT,
            exchange_delivery_days TEXT,
            exchange_sent_at TEXT,
            exchange_received_at TEXT,
            exchange_customer_acceptance TEXT,
            exchange_round INTEGER DEFAULT 1
        )
    """)

    # Migration for databases made by previous versions.
    for table, fields in {
        "orders": [
            ("chat_id", "TEXT"), ("created_at", "TEXT"), ("delivered_at", "TEXT"),
            ("accepted_at", "TEXT"), ("customer_acceptance", "TEXT"),
            ("website_order_id", "INTEGER"), ("website_customer_ref", "TEXT"),
            ("province", "TEXT"), ("city", "TEXT")
        ],
        "products": [
            ("image_file_id", "TEXT"), ("category_id", "INTEGER"),
            ("created_at", "TEXT")
        ],
        "return_requests": [
            ("return_deadline", "TEXT"), ("customer_name", "TEXT"), ("phone", "TEXT"),
            ("exchange_product", "TEXT"), ("exchange_difference", "TEXT"),
            ("refund_amount", "TEXT"), ("bank_account_name", "TEXT"),
            ("card_number", "TEXT"), ("iban", "TEXT"), ("bank_name", "TEXT"),
            ("paid_at", "TEXT"), ("exchange_delivery_days", "TEXT"),
            ("exchange_sent_at", "TEXT"), ("exchange_received_at", "TEXT"),
            ("exchange_customer_acceptance", "TEXT"), ("exchange_round", "INTEGER DEFAULT 1")
        ]
    }.items():
        for name, definition in fields:
            ensure_column(conn, table, name, definition)

    # Default category structure. These are only created once.
    defaults = [
        (None, "پوشاک"),
        (None, "کفش"),
        (None, "اکسسوری و زیورآلات"),
        (None, "طلا"),
        (None, "عطر و ادکلن"),
    ]
    for parent, name in defaults:
        exists = conn.execute(
            "SELECT id FROM categories WHERE parent_id IS NULL AND name=?", (name,)
        ).fetchone()
        if not exists:
            conn.execute(
                "INSERT INTO categories(parent_id,name,active,sort_order,created_at) VALUES(?,?,?,?,?)",
                (parent, name, 1, 0, now_string())
            )

    clothing = conn.execute(
        "SELECT id FROM categories WHERE parent_id IS NULL AND name='پوشاک'"
    ).fetchone()
    shoes = conn.execute(
        "SELECT id FROM categories WHERE parent_id IS NULL AND name='کفش'"
    ).fetchone()
    perfume = conn.execute(
        "SELECT id FROM categories WHERE parent_id IS NULL AND name='عطر و ادکلن'"
    ).fetchone()
    if clothing:
        for name in ("مردانه", "زنانه", "بچگانه"):
            if not conn.execute("SELECT id FROM categories WHERE parent_id=? AND name=?", (clothing[0], name)).fetchone():
                conn.execute("INSERT INTO categories(parent_id,name,active,created_at) VALUES(?,?,1,?)", (clothing[0], name, now_string()))
    if shoes:
        for name in ("مردانه", "زنانه", "بچگانه"):
            if not conn.execute("SELECT id FROM categories WHERE parent_id=? AND name=?", (shoes[0], name)).fetchone():
                conn.execute("INSERT INTO categories(parent_id,name,active,created_at) VALUES(?,?,1,?)", (shoes[0], name, now_string()))
    if perfume:
        for name in ("مردانه", "زنانه", "یونیسکس"):
            if not conn.execute("SELECT id FROM categories WHERE parent_id=? AND name=?", (perfume[0], name)).fetchone():
                conn.execute("INSERT INTO categories(parent_id,name,active,created_at) VALUES(?,?,1,?)", (perfume[0], name, now_string()))

    # Keep P-001 from the old system so existing customers/orders remain compatible.
    if not conn.execute("SELECT id FROM products WHERE code='P-001'").fetchone():
        conn.execute("""
            INSERT INTO products(code,name,description,price,delivery_days,image_path,active,created_at)
            VALUES(?,?,?,?,?,?,?,?)
        """, (
            "P-001", "محصول نمونه", "توضیحات محصول را اینجا وارد کنید.",
            "1500000", "1 تا 3", PRODUCT_IMAGE_PATH, 1, now_string()
        ))

    conn.commit()
    conn.close()


# ============================================================
# COMMON HELPERS
# ============================================================
def get_user_id(message):
    return getattr(message, "sender_id", None)


def get_chat_id(message):
    return getattr(message, "chat_id", None)


def is_admin(message):
    return get_chat_id(message) == ADMIN_CHAT_ID


def button_id(message):
    try:
        return message.aux_data.button_id
    except Exception:
        return None


def make_customer_ref(chat_id):
    raw = str(chat_id or "").encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def website_order_link(chat_id, product_code):
    ref = make_customer_ref(chat_id)
    return f"{WEBSITE_BASE_URL}landing.html?product={product_code}&customer_ref={ref}"


def deadline(accepted, hours=RETURN_WINDOW_HOURS):
    dt = parse_dt(accepted)
    return dt + timedelta(hours=hours) if dt else None


def status_fa(status):
    return {
        "جدید": "🆕 جدید",
        "تأیید شده": "✅ تأیید شده",
        "لغو شده": "❌ لغو شده",
        "تحویل شده": "🚚 تحویل شده",
        "تحویل و تأیید شده": "📦 تحویل و تأیید شده",
        "تعویض انجام شد": "🔄 تعویض انجام شد",
        "مرجوع و وجه بازگشت داده شد": "💰 مرجوع و وجه بازگشت داده شد",
    }.get(status, status or "نامشخص")


def after_status(status):
    return {
        "در انتظار بررسی": "🟡 در انتظار بررسی",
        "در انتظار بررسی مجدد": "🟡 در انتظار بررسی مجدد",
        "تأیید شده": "✅ تأیید شده",
        "رد شده": "❌ رد شده",
        "در حال تعویض": "🔄 در حال تعویض",
        "کالای تعویضی ارسال شد": "📦 کالای تعویضی ارسال شد",
        "در انتظار تأیید دریافت تعویضی": "🕐 در انتظار تأیید دریافت تعویضی",
        "تعویض انجام شد": "✅ تعویض انجام شد",
        "در انتظار واریز": "💰 در انتظار واریز",
        "وجه واریز شد": "✅ وجه واریز شد",
        "تکمیل شده": "✅ تکمیل شده",
    }.get(status, status or "نامشخص")


def get_product(code):
    conn = get_db()
    row = conn.execute("""
        SELECT p.*, c.name AS category_name
        FROM products p
        LEFT JOIN categories c ON c.id=p.category_id
        WHERE p.code=? AND p.active=1
    """, (code,)).fetchone()
    conn.close()
    return row


def get_admin_product(code):
    conn = get_db()
    row = conn.execute("SELECT * FROM products WHERE code=?", (code,)).fetchone()
    conn.close()
    return row


def get_product_stats(code):
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) FROM orders WHERE product=?", (code,)).fetchone()[0]
    avg, count = conn.execute("SELECT AVG(rating),COUNT(*) FROM reviews WHERE product=?", (code,)).fetchone()
    conn.close()
    return total, round(avg or 0, 1), count

# ============================================================
# KEYPADS
# ============================================================
def main_menu():
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id="show_categories", text="🗂 مشاهده دسته‌بندی محصولات"))
        .row(ChatKeypadBuilder().button(id="my_orders", text="📋 سفارش‌های من"),
             ChatKeypadBuilder().button(id="shop_policy", text="📜 شرایط فروش و خدمات پس از فروش"))
        .build())


def product_menu(code):
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id=f"order_product_{code}", text="🛒 ادامه خرید این محصول"))
        .row(ChatKeypadBuilder().button(id="show_categories", text="🗂 سایر محصولات"))
        .row(ChatKeypadBuilder().button(id="main_menu", text="🏠 منوی اصلی"))
        .build())


def category_menu(parent_id=0):
    conn = get_db()
    if parent_id:
        rows = conn.execute("SELECT id,name FROM categories WHERE parent_id=? AND active=1 ORDER BY sort_order,id", (parent_id,)).fetchall()
    else:
        rows = conn.execute("SELECT id,name FROM categories WHERE parent_id IS NULL AND active=1 ORDER BY sort_order,id").fetchall()
    conn.close()
    builder = ChatKeypadBuilder()
    for row in rows:
        builder.row(ChatKeypadBuilder().button(id=f"cat_{row['id']}", text=f"📂 {row['name']}"))
    if parent_id:
        builder.row(ChatKeypadBuilder().button(id="show_categories", text="↩️ دسته‌بندی‌های اصلی"))
    builder.row(ChatKeypadBuilder().button(id="main_menu", text="🏠 منوی اصلی"))
    return builder.build()


def products_in_category_menu(category_id):
    conn = get_db()
    rows = conn.execute("SELECT code,name FROM products WHERE category_id=? AND active=1 ORDER BY id DESC", (category_id,)).fetchall()
    parent = conn.execute("SELECT parent_id FROM categories WHERE id=?", (category_id,)).fetchone()
    conn.close()
    builder = ChatKeypadBuilder()
    for row in rows:
        builder.row(ChatKeypadBuilder().button(id=f"view_product_{row['code']}", text=f"📦 {row['code']} | {row['name']}"))
    if not rows:
        # The caller sends the explanatory message separately.
        pass
    if parent and parent[0]:
        builder.row(ChatKeypadBuilder().button(id=f"cat_{parent[0]}", text="↩️ زیر‌دسته قبلی"))
    else:
        builder.row(ChatKeypadBuilder().button(id="show_categories", text="↩️ دسته‌بندی‌ها"))
    builder.row(ChatKeypadBuilder().button(id="main_menu", text="🏠 منوی اصلی"))
    return builder.build(), rows


def confirm_menu():
    return ChatKeypadBuilder().row(
        ChatKeypadBuilder().button(id="customer_confirm_order", text="✅ تأیید سفارش"),
        ChatKeypadBuilder().button(id="customer_cancel_order", text="❌ لغو سفارش")
    ).build()


def accept_menu(oid):
    return ChatKeypadBuilder().row(
        ChatKeypadBuilder().button(id=f"customer_accept_delivery_{oid}", text="✅ کالا را رویت کردم و تحویل گرفتم")
    ).build()


def order_admin_menu(oid):
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id=f"admin_confirm_{oid}", text="✅ تأیید سفارش"),
             ChatKeypadBuilder().button(id=f"admin_cancel_{oid}", text="❌ لغو سفارش"))
        .row(ChatKeypadBuilder().button(id=f"admin_deliver_{oid}", text="🚚 ثبت تحویل"))
        .build())


def customer_after_menu(oid, exchange=False):
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id=f"customer_rate_{oid}", text="⭐ رضایت دارم؛ امتیاز و نظر"))
        .row(ChatKeypadBuilder().button(id=f"customer_exchange_{oid}", text="🔄 مجدداً درخواست تعویض" if exchange else "🔄 درخواست تعویض"),
             ChatKeypadBuilder().button(id=f"customer_refund_{oid}", text="💰 مرجوعی و بازگشت وجه"))
        .build())


def exchange_accept_menu(rid):
    return ChatKeypadBuilder().row(ChatKeypadBuilder().button(id=f"customer_accept_exchange_{rid}", text="✅ کالای تعویضی را دریافت کردم")).build()


def refund_accept_menu(rid):
    return ChatKeypadBuilder().row(ChatKeypadBuilder().button(id=f"customer_accept_refund_{rid}", text="✅ دریافت وجه را تأیید می‌کنم")).build()


def admin_menu():
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id="admin_customer_panel", text="👤 امور مشتریان"),
             ChatKeypadBuilder().button(id="admin_categories", text="🗂 دسته‌بندی‌ها"))
        .row(ChatKeypadBuilder().button(id="admin_products", text="📦 مدیریت محصولات"))
        .row(ChatKeypadBuilder().button(id="admin_export_orders", text="📊 گزارش Excel"))
        .build())


def admin_customer_menu():
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id="admin_orders", text="📦 سفارش‌ها"),
             ChatKeypadBuilder().button(id="admin_after_sales", text="🔄 خدمات پس از فروش"))
        .row(ChatKeypadBuilder().button(id="admin_export_orders", text="📊 گزارش Excel"))
        .row(ChatKeypadBuilder().button(id="admin_back", text="🏠 بازگشت به صفحه اصلی"))
        .build())


def admin_product_category_menu():
    """انتخاب دسته‌ای که قبلاً در بخش «مدیریت دسته‌بندی‌ها» تعریف شده است."""
    conn = get_db()
    roots = conn.execute(
        "SELECT id,name FROM categories WHERE parent_id IS NULL AND active=1 ORDER BY sort_order,id"
    ).fetchall()
    children = conn.execute(
        "SELECT c.id,c.name,c.parent_id,p.name AS parent_name "
        "FROM categories c JOIN categories p ON p.id=c.parent_id "
        "WHERE c.active=1 AND p.active=1 ORDER BY p.sort_order,p.id,c.sort_order,c.id"
    ).fetchall()
    conn.close()

    b = ChatKeypadBuilder()
    if not roots:
        b.row(ChatKeypadBuilder().button(id="admin_categories", text="🗂 تعریف دسته‌بندی‌ها"))
    else:
        for root in roots:
            b.row(ChatKeypadBuilder().button(
                id=f"admin_product_add_cat_{root['id']}",
                text=f"📂 {root['name']}"
            ))
            for child in [c for c in children if c['parent_id'] == root['id']]:
                b.row(ChatKeypadBuilder().button(
                    id=f"admin_product_add_cat_{child['id']}",
                    text=f"↳ {child['name']} | {root['name']}"
                ))

    b.row(ChatKeypadBuilder().button(id="admin_categories", text="🗂 مدیریت دسته‌بندی‌ها"))
    b.row(ChatKeypadBuilder().button(id="admin_products", text="↩️ مدیریت محصولات"),
          ChatKeypadBuilder().button(id="admin_back", text="🏠 خانه"))
    return b.build()


def admin_categories_menu():
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id="admin_category_add", text="➕ افزودن دسته اصلی"))
        .row(ChatKeypadBuilder().button(id="admin_subcategory_add", text="➕ افزودن زیر‌دسته"))
        .row(ChatKeypadBuilder().button(id="admin_category_list", text="📋 مشاهده دسته‌ها"))
        .row(ChatKeypadBuilder().button(id="admin_back", text="🏠 بازگشت به صفحه اصلی"))
        .build())


def admin_products_menu():
    return (ChatKeypadBuilder()
        .row(ChatKeypadBuilder().button(id="admin_product_add", text="➕ افزودن محصول"))
        .row(ChatKeypadBuilder().button(id="admin_product_list", text="📋 محصولات فعال"))
        .row(ChatKeypadBuilder().button(id="admin_inactive_products", text="🚫 محصولات غیرفعال"))
        .row(ChatKeypadBuilder().button(id="admin_edit_product_select", text="✏️ ویرایش محصول"),
             ChatKeypadBuilder().button(id="admin_delete_product_select", text="🗑 حذف محصول"))
        .row(ChatKeypadBuilder().button(id="admin_back", text="🏠 بازگشت به صفحه اصلی"))
        .build())


def return_admin_menu(rid, typ, status="در انتظار بررسی"):
    b = ChatKeypadBuilder()
    if status in ("در انتظار بررسی", "در انتظار بررسی مجدد"):
        b.row(ChatKeypadBuilder().button(id=f"admin_after_approve_{rid}", text="✅ تأیید درخواست"),
              ChatKeypadBuilder().button(id=f"admin_after_reject_{rid}", text="❌ رد درخواست"))
    elif typ == "مرجوعی و بازگشت وجه" and status == "در انتظار واریز":
        b.row(ChatKeypadBuilder().button(id=f"admin_refund_paid_{rid}", text="💰 ثبت واریز وجه"))
    elif typ == "تعویض" and status == "در حال تعویض":
        b.row(ChatKeypadBuilder().button(id=f"admin_exchange_sent_{rid}", text="📦 ثبت ارسال کالای تعویضی"))
    b.row(ChatKeypadBuilder().button(id="admin_customer_panel", text="↩️ امور مشتریان"))
    b.row(ChatKeypadBuilder().button(id="admin_back", text="🏠 بازگشت به صفحه اصلی"))
    return b.build()


def rating_menu(oid):
    return (ChatKeypadBuilder()
        .row(*[ChatKeypadBuilder().button(id=f"rating_{oid}_{i}", text=f"⭐ {i}") for i in (1, 2, 3)])
        .row(ChatKeypadBuilder().button(id=f"rating_{oid}_4", text="⭐ 4"),
             ChatKeypadBuilder().button(id=f"rating_{oid}_5", text="⭐ 5"))
        .build())

# ============================================================
# PRODUCT DISPLAY
# ============================================================
def build_product_text(code):
    p = get_product(code)
    if not p:
        return "❌ محصول موردنظر پیدا نشد."
    total, avg, review_count = get_product_stats(code)
    conn = get_db()
    recent_since = (datetime.now() - timedelta(days=RECENT_ORDER_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    recent = conn.execute("SELECT COUNT(*) FROM orders WHERE product=? AND created_at>=?", (code, recent_since)).fetchone()[0]
    conn.close()

    text = (
        f"🏪 سبلان شاپ | از دل آذربایجان تا خانه شما\n\n"
        f"📦 کد محصول: {p['code']}\n"
        f"🔹 نام محصول: {p['name']}\n\n"
        f"📝 توضیحات:\n{p['description'] or '-'}\n\n"
        f"💰 قیمت: {price(p['price'])} تومان\n\n"
        f"🛒 تعداد سفارش ثبت‌شده: {total}\n"
    )
    if recent >= RECENT_ORDER_THRESHOLD:
        text += "🔥 این محصول اخیراً مورد توجه مشتریان بوده است.\n"
    text += (
        f"\n🚚 ارسال رایگان فقط در محدوده شهر {FREE_DELIVERY_CITY}\n"
        "📦 ارسال به سایر نقاط کشور با احتساب هزینه پستی\n"
        f"⏱ زمان تقریبی ارسال: {p['delivery_days'] or '-'} روز کاری\n"
        "🏠 پرداخت درب منزل پس از رویت و تحویل کالا\n\n"
        f"⭐ امتیاز مشتریان: {avg} / 5\n"
        f"💬 تعداد نظر ثبت‌شده: {review_count}\n\n"
        "برای ادامه خرید از دکمه زیر استفاده کنید."
    )
    return text


async def show_product(chat_id, code):
    p = get_product(code)
    if not p:
        return await bot.send_message(chat_id=chat_id, text="❌ محصول موجود نیست.", chat_keypad=main_menu(), chat_keypad_type="New")
    text = build_product_text(code)
    path = p["image_path"]
    fid = p["image_file_id"]
    keypad = product_menu(code)
    if path and os.path.exists(path):
        try:
            return await bot.send_image(chat_id=chat_id, path=path, text=text, chat_keypad=keypad, chat_keypad_type="New")
        except Exception as exc:
            print("IMAGE ERROR:", exc)
    if fid:
        try:
            return await bot.send_image(chat_id=chat_id, file_id=fid, text=text, chat_keypad=keypad, chat_keypad_type="New")
        except Exception as exc:
            print("FILE IMAGE ERROR:", exc)
    return await bot.send_message(chat_id=chat_id, text="⚠️ تصویر محصول ثبت نشده است.\n\n" + text, chat_keypad=keypad, chat_keypad_type="New")


def policy_text():
    return (
        "📜 شرایط فروش و خدمات پس از فروش\n\n"
        "🏠 پرداخت درب منزل پس از رویت و تحویل کالا.\n\n"
        f"🚚 ارسال رایگان فقط در محدوده شهر {FREE_DELIVERY_CITY}.\n\n"
        "📦 ارسال به سایر نقاط کشور با احتساب هزینه پستی.\n\n"
        f"1️⃣ مهلت عادی تعویض/مرجوعی: {RETURN_WINDOW_HOURS} ساعت پس از تأیید دریافت.\n\n"
        f"2️⃣ مهلت درخواست تعویض: {EXCHANGE_WINDOW_HOURS} ساعت پس از تأیید دریافت.\n\n"
        "3️⃣ درخواست توسط مدیر بررسی و نتیجه در روبیکا به مشتری اعلام می‌شود.\n\n"
        "4️⃣ در صورت مرجوعی، اطلاعات بانکی دریافت و پس از تأیید مدیر، وجه واریز می‌شود."
    )

# ============================================================
# NOTIFICATIONS / WEBSITE SYNC
# ============================================================
async def notify_customer_status(oid, status):
    conn = get_db()
    row = conn.execute("SELECT chat_id,name,product FROM orders WHERE id=?", (oid,)).fetchone()
    conn.close()
    if row and row["chat_id"]:
        await bot.send_message(
            chat_id=row["chat_id"],
            text=f"📢 وضعیت سفارش #{oid} تغییر کرد.\n\n📦 محصول: {row['product']}\n📌 وضعیت: {status_fa(status)}"
        )


async def notify_admin_order(oid):
    conn = get_db()
    row = conn.execute("SELECT name,phone,address,product,status FROM orders WHERE id=?", (oid,)).fetchone()
    conn.close()
    if not row:
        return
    await bot.send_message(
        chat_id=ADMIN_CHAT_ID,
        text=(f"🆕 سفارش جدید #{oid}\n\n👤 نام: {row['name']}\n📱 تلفن: {row['phone']}\n"
              f"📍 آدرس: {row['address']}\n📦 محصول: {row['product']}\n📌 وضعیت: {status_fa(row['status'])}"),
        chat_keypad=order_admin_menu(oid),
        chat_keypad_type="New"
    )


def website_sync_worker():
    print("WEBSITE ORDER SYNC STARTED")
    while True:
        try:
            with urllib.request.urlopen(WEBSITE_API_URL, timeout=5) as response:
                data = json.loads(response.read().decode("utf-8"))
            orders = data.get("orders", []) if isinstance(data, dict) else []
            for w in orders:
                wid = w.get("id")
                if wid is None:
                    continue
                conn = get_db()
                if conn.execute("SELECT id FROM orders WHERE website_order_id=?", (int(wid),)).fetchone():
                    conn.close()
                    continue

                product_code = str(w.get("product_code") or "-")
                product = conn.execute("SELECT name FROM products WHERE code=?", (product_code,)).fetchone()
                product_name = f"{product_code} | {product['name']}" if product else product_code

                ref = str(w.get("customer_ref") or "")
                chat_id = ""
                if ref:
                    try:
                        padded = ref + "=" * (-len(ref) % 4)
                        chat_id = base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8")
                    except Exception:
                        chat_id = ""

                address = f"{w.get('province') or '-'}، {w.get('city') or '-'}، {w.get('address') or '-'}"
                cur = conn.execute("""
                    INSERT INTO orders(
                        website_order_id,website_customer_ref,user_id,product,name,phone,province,city,address,status,chat_id,created_at
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    int(wid), ref, f"website:{wid}", product_name,
                    str(w.get("fullname") or "-"), str(w.get("phone") or "-"),
                    str(w.get("province") or "-"), str(w.get("city") or "-"), address,
                    str(w.get("status") or "جدید"), chat_id,
                    str(w.get("created_at") or now_string())
                ))
                oid = cur.lastrowid
                conn.commit()
                conn.close()
                print(f"WEBSITE ORDER IMPORTED: website #{wid} -> bot #{oid}")

                try:
                    loop = BOT_LOOP
                    if loop and not loop.is_closed():
                        future = asyncio.run_coroutine_threadsafe(notify_admin_order(oid), loop)
                        future.result(timeout=15)
                except Exception as exc:
                    print("WEBSITE ADMIN NOTIFICATION ERROR:", exc)

        except Exception as exc:
            print("WEBSITE SYNC ERROR:", exc)
        import time
        time.sleep(WEBSITE_SYNC_INTERVAL)

# ============================================================
# AFTER SALES
# ============================================================
async def notify_after_sales(rid):
    conn = get_db()
    row = conn.execute("""
        SELECT r.order_id,r.request_type,r.reason,r.status,o.name,o.phone,o.address,o.product,
               r.refund_amount,r.exchange_product,r.card_number,r.iban,r.bank_name,r.exchange_delivery_days
        FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=?
    """, (rid,)).fetchone()
    conn.close()
    if not row:
        return False
    text = (
        f"🔔 درخواست خدمات پس از فروش #{rid}\n\n"
        f"🧾 سفارش: #{row['order_id']}\n👤 مشتری: {row['name']}\n📱 تلفن: {row['phone']}\n"
        f"📍 آدرس: {row['address']}\n📦 محصول: {row['product']}\n"
        f"🔄 نوع درخواست: {row['request_type']}\n📝 دلیل: {row['reason']}\n"
        f"💰 مبلغ: {price(row['refund_amount'])} تومان\n💳 کارت: {row['card_number'] or '-'}\n"
        f"🏦 شبا: {row['iban'] or '-'}\n🏦 بانک: {row['bank_name'] or '-'}\n"
        f"📌 وضعیت: {after_status(row['status'])}"
    )
    try:
        await bot.send_message(chat_id=ADMIN_CHAT_ID, text=text)
        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=f"🎛 مدیریت درخواست #{rid}\nوضعیت: {after_status(row['status'])}",
            chat_keypad=return_admin_menu(rid, row['request_type'], row['status']),
            chat_keypad_type="New"
        )
        return True
    except Exception as exc:
        print("AFTER SALES NOTIFICATION ERROR:", exc)
        return False


async def notify_customer_after(rid):
    conn = get_db()
    row = conn.execute("""
        SELECT r.order_id,r.request_type,r.status,r.admin_note,o.chat_id,r.exchange_delivery_days
        FROM return_requests r JOIN orders o ON o.id=r.order_id WHERE r.id=?
    """, (rid,)).fetchone()
    conn.close()
    if not row or not row["chat_id"]:
        return
    text = (
        f"📢 وضعیت درخواست #{rid} برای سفارش #{row['order_id']}\n\n"
        f"🔄 نوع: {row['request_type']}\n📌 وضعیت: {after_status(row['status'])}\n"
        f"📝 توضیح: {row['admin_note'] or '-'}"
    )
    if row["request_type"] == "تعویض" and row["status"] == "در حال تعویض":
        text += f"\n\n📦 درخواست تعویض تأیید شد.\n⏱ زمان تقریبی ارسال: {row['exchange_delivery_days'] or '-'} روز کاری."
    elif row["request_type"] == "مرجوعی و بازگشت وجه" and row["status"] == "در انتظار واریز":
        text += "\n\n💰 درخواست مرجوعی تأیید شد. پس از واریز واقعی، پیام تأیید دریافت وجه برای شما ارسال می‌شود."
    elif row["request_type"] == "مرجوعی و بازگشت وجه" and row["status"] == "وجه واریز شد":
        text += "\n\n💰 وجه واریز شده است. لطفاً دریافت وجه را تأیید کنید."
    elif row["request_type"] == "تعویض" and row["status"] == "کالای تعویضی ارسال شد":
        text += "\n\n📦 کالای تعویضی ارسال شده است. لطفاً پس از دریافت، آن را تأیید کنید."

    if row["request_type"] == "مرجوعی و بازگشت وجه" and row["status"] == "وجه واریز شد":
        await bot.send_message(chat_id=row["chat_id"], text=text, chat_keypad=refund_accept_menu(rid), chat_keypad_type="New")
    elif row["request_type"] == "تعویض" and row["status"] == "کالای تعویضی ارسال شد":
        await bot.send_message(chat_id=row["chat_id"], text=text, chat_keypad=exchange_accept_menu(rid), chat_keypad_type="New")
    else:
        await bot.send_message(chat_id=row["chat_id"], text=text)

# ============================================================
# EXCEL
# ============================================================
async def export_excel(chat_id):
    if not OPENPYXL_AVAILABLE:
        return await bot.send_message(chat_id=chat_id, text="❌ openpyxl نصب نیست.\n\npip install openpyxl")
    try:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "orders_export.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append([
            "شماره سفارش", "شناسه کاربر", "شناسه روبیکا", "محصول", "نام مشتری", "تلفن",
            "استان", "شهر", "آدرس", "وضعیت", "تاریخ ثبت", "تاریخ تحویل", "تأیید دریافت", "امتیاز", "نظر"
        ])
        conn = get_db()
        rows = conn.execute("""
            SELECT o.id,o.user_id,o.chat_id,o.product,o.name,o.phone,o.province,o.city,o.address,
                   o.status,o.created_at,o.delivered_at,o.customer_acceptance,rv.rating,rv.comment
            FROM orders o LEFT JOIN reviews rv ON rv.order_id=o.id ORDER BY o.id DESC
        """).fetchall()
        for row in rows:
            ws.append(list(row))

        ws2 = wb.create_sheet("AfterSales")
        ws2.append([
            "شناسه درخواست", "شماره سفارش", "شناسه کاربر", "نوع", "دلیل", "وضعیت",
            "مبلغ بازگشت", "کارت", "شبا", "بانک", "کالای جایگزین", "تاریخ ثبت", "تصمیم", "واریز"
        ])
        rows2 = conn.execute("""
            SELECT id,order_id,user_id,request_type,reason,status,refund_amount,card_number,iban,
                   bank_name,exchange_product,created_at,decided_at,paid_at
            FROM return_requests ORDER BY id DESC
        """).fetchall()
        for row in rows2:
            ws2.append(list(row))

        ws3 = wb.create_sheet("Products")
        ws3.append(["کد", "نام", "دسته", "قیمت", "زمان ارسال", "فعال", "تعداد سفارش", "امتیاز", "تعداد نظر"])
        products = conn.execute("""
            SELECT p.code,p.name,c.name,p.price,p.delivery_days,p.active,
                   (SELECT COUNT(*) FROM orders o WHERE o.product=p.code),
                   (SELECT ROUND(AVG(r.rating),1) FROM reviews r WHERE r.product=p.code),
                   (SELECT COUNT(*) FROM reviews r2 WHERE r2.product=p.code)
            FROM products p LEFT JOIN categories c ON c.id=p.category_id ORDER BY p.id DESC
        """).fetchall()
        for row in products:
            ws3.append(list(row))
        conn.close()

        for sheet in wb.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in sheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            for column in sheet.columns:
                letter = column[0].column_letter
                max_len = max(len(str(c.value or "")) for c in column)
                sheet.column_dimensions[letter].width = min(max(max_len + 3, 12), 40)

        wb.save(path)
        await bot.send_document(chat_id=chat_id, path=path, text="📊 گزارش کامل فروشگاه سبلان شاپ", file_name="orders_export.xlsx")
        print("EXCEL REPORT SENT:", path)
    except Exception as exc:
        print("EXCEL ERROR:", exc)
        await bot.send_message(chat_id=chat_id, text=f"❌ خطا در تهیه گزارش Excel:\n{type(exc).__name__}: {exc}")

# ============================================================
# ADMIN STATES
# ============================================================
def category_list_text():
    conn = get_db()
    roots = conn.execute("SELECT id,name,active FROM categories WHERE parent_id IS NULL ORDER BY id").fetchall()
    lines = ["🗂 دسته‌بندی‌های فروشگاه\n"]
    for root in roots:
        lines.append(f"{'🟢' if root['active'] else '🔴'} {root['id']} - {root['name']}")
        children = conn.execute("SELECT id,name,active FROM categories WHERE parent_id=? ORDER BY id", (root['id'],)).fetchall()
        for child in children:
            lines.append(f"   └─ {'🟢' if child['active'] else '🔴'} {child['id']} - {child['name']}")
    conn.close()
    return "\n".join(lines)


async def admin_product_list(chat_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT p.code,p.name,p.price,p.active,c.name AS category
        FROM products p LEFT JOIN categories c ON c.id=p.category_id
        WHERE p.active=1 ORDER BY p.id DESC
    """).fetchall()
    conn.close()
    if not rows:
        return await bot.send_message(
            chat_id=chat_id,
            text="📦 در حال حاضر محصول فعالی وجود ندارد.",
            chat_keypad=admin_products_menu(),
            chat_keypad_type="New"
        )
    await bot.send_message(chat_id=chat_id, text="📋 محصولات فعال")
    for row in rows:
        keypad = (ChatKeypadBuilder()
            .row(
                ChatKeypadBuilder().button(id=f"admin_toggle_product_{row['code']}", text="🚫 غیرفعال کردن")
            )
            .row(
                ChatKeypadBuilder().button(id=f"admin_edit_product_{row['code']}", text="✏️ ویرایش")
            )
            .build())
        await bot.send_message(
            chat_id=chat_id,
            text=(f"📦 {row['code']}\n"
                  f"نام: {row['name']}\n"
                  f"دسته: {row['category'] or '-'}\n"
                  f"قیمت: {price(row['price'])} تومان\n"
                  "وضعیت: 🟢 فعال"),
            chat_keypad=keypad,
            chat_keypad_type="New"
        )
    return await bot.send_message(
        chat_id=chat_id,
        text="از منوی مدیریت محصولات می‌توانید محصولی را ویرایش، حذف یا غیرفعال کنید.",
        chat_keypad=admin_products_menu(),
        chat_keypad_type="New"
    )


async def admin_inactive_product_list(chat_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT p.code,p.name,p.price,p.active,c.name AS category
        FROM products p LEFT JOIN categories c ON c.id=p.category_id
        WHERE p.active=0 ORDER BY p.id DESC
    """).fetchall()
    conn.close()
    if not rows:
        return await bot.send_message(
            chat_id=chat_id,
            text="🚫 محصول غیرفعالی وجود ندارد.",
            chat_keypad=admin_products_menu(),
            chat_keypad_type="New"
        )
    await bot.send_message(chat_id=chat_id, text="🚫 محصولات غیرفعال")
    for row in rows:
        keypad = (ChatKeypadBuilder()
            .row(ChatKeypadBuilder().button(id=f"admin_toggle_product_{row['code']}", text="♻️ فعال‌سازی"))
            .row(ChatKeypadBuilder().button(id=f"admin_edit_product_{row['code']}", text="✏️ ویرایش"),
                 ChatKeypadBuilder().button(id=f"admin_delete_product_{row['code']}", text="🗑 حذف"))
            .build())
        await bot.send_message(
            chat_id=chat_id,
            text=(f"📦 {row['code']}\n"
                  f"نام: {row['name']}\n"
                  f"دسته: {row['category'] or '-'}\n"
                  f"قیمت: {price(row['price'])} تومان\n"
                  "وضعیت: 🔴 غیرفعال"),
            chat_keypad=keypad,
            chat_keypad_type="New"
        )
    return await bot.send_message(
        chat_id=chat_id,
        text="برای بازگشت، از دکمه مدیریت محصولات استفاده کنید.",
        chat_keypad=admin_products_menu(),
        chat_keypad_type="New"
    )


async def admin_product_selection(chat_id, action):
    conn = get_db()
    rows = conn.execute("SELECT code,name,active FROM products ORDER BY id ASC").fetchall()
    conn.close()
    if not rows:
        return await bot.send_message(chat_id=chat_id, text="📦 هیچ محصولی وجود ندارد.", chat_keypad=admin_products_menu(), chat_keypad_type="New")
    title = "✏️ کدام محصول را می‌خواهید ویرایش کنید؟" if action == "edit" else "🗑 کدام محصول را می‌خواهید حذف کنید؟"
    builder = ChatKeypadBuilder()
    for i, row in enumerate(rows, 1):
        status = "🟢" if row['active'] else "🔴"
        bid = f"admin_select_edit_{row['code']}" if action == "edit" else f"admin_select_delete_{row['code']}"
        builder.row(ChatKeypadBuilder().button(id=bid, text=f"{i}️⃣ {status} {row['code']} | {row['name']}"))
    builder.row(ChatKeypadBuilder().button(id="admin_products", text="↩️ مدیریت محصولات"),
                ChatKeypadBuilder().button(id="admin_back", text="🏠 خانه"))
    return await bot.send_message(chat_id=chat_id, text=title, chat_keypad=builder.build(), chat_keypad_type="New")


# ============================================================
# MESSAGE HANDLER
# ============================================================
@bot.on_message()
async def handle_message(bot_instance: Robot, message: Message):
    global BOT_LOOP
    try:
        BOT_LOOP = asyncio.get_running_loop()
    except RuntimeError:
        pass

    uid = get_user_id(message)
    cid = get_chat_id(message)
    text = (getattr(message, "text", "") or "").strip()

    if text == "/start":
        if cid == ADMIN_CHAT_ID:
            return await bot.send_message(chat_id=cid, text="👨‍💼 پنل مدیریت سبلان شاپ", chat_keypad=admin_menu(), chat_keypad_type="New")
        return await bot.send_message(chat_id=cid, text="🏪 سبلان شاپ\n\nلطفاً گزینه موردنظر را انتخاب کنید.", chat_keypad=main_menu(), chat_keypad_type="New")

    # ---------------- ADMIN INPUT STATES ----------------
    if cid == ADMIN_CHAT_ID and cid in admin_states:
        state = admin_states[cid]

        if text == "لغو":
            admin_states.pop(cid, None)
            return await bot.send_message(chat_id=cid, text="❌ عملیات لغو شد.", chat_keypad=admin_menu(), chat_keypad_type="New")

        if state == "category_name":
            if not text:
                return await bot.send_message(chat_id=cid, text="❌ نام دسته نمی‌تواند خالی باشد.")
            conn = get_db()
            conn.execute("INSERT INTO categories(parent_id,name,active,created_at) VALUES(NULL,?,1,?)", (text, now_string()))
            conn.commit(); conn.close(); admin_states.pop(cid, None)
            return await bot.send_message(chat_id=cid, text="✅ دسته اصلی اضافه شد.", chat_keypad=admin_categories_menu(), chat_keypad_type="New")

        if state == "subcategory_parent":
            if not norm(text).isdigit():
                return await bot.send_message(chat_id=cid, text="❌ شناسه دسته اصلی را به صورت عدد وارد کنید.")
            parent = int(norm(text))
            conn = get_db(); row = conn.execute("SELECT id,name FROM categories WHERE id=? AND parent_id IS NULL", (parent,)).fetchone(); conn.close()
            if not row:
                return await bot.send_message(chat_id=cid, text="❌ دسته اصلی پیدا نشد.")
            admin_states[cid] = {"state": "subcategory_name", "parent_id": parent}
            return await bot.send_message(chat_id=cid, text=f"📝 نام زیر‌دسته برای «{row['name']}» را وارد کنید:")

        if isinstance(state, dict) and state.get("state") == "subcategory_name":
            parent = state["parent_id"]
            conn = get_db(); conn.execute("INSERT INTO categories(parent_id,name,active,created_at) VALUES(?,?,1,?)", (parent,text,now_string())); conn.commit(); conn.close()
            admin_states.pop(cid, None)
            return await bot.send_message(chat_id=cid, text="✅ زیر‌دسته اضافه شد.", chat_keypad=admin_categories_menu(), chat_keypad_type="New")

        if state == "product_add_new_root_name":
            if not text:
                return await bot.send_message(chat_id=cid, text="❌ نام دسته اصلی نمی‌تواند خالی باشد.")
            base = admin_states[cid]
            conn = get_db()
            exists = conn.execute("SELECT id FROM categories WHERE parent_id IS NULL AND name=?", (text,)).fetchone()
            if exists:
                conn.close()
                return await bot.send_message(chat_id=cid, text="❌ این دسته اصلی قبلاً وجود دارد. از همان دسته قبلی استفاده کنید.", chat_keypad=admin_product_category_menu(), chat_keypad_type="New")
            cur = conn.execute("INSERT INTO categories(parent_id,name,active,sort_order,created_at) VALUES(NULL,?,1,0,?)", (text, now_string()))
            parent_id = cur.lastrowid
            conn.commit(); conn.close()
            # بعد از ساخت دسته اصلی، حتماً نام زیر‌دسته را می‌گیریم.
            admin_states[cid] = {
                "state":"product_add_new_root_sub_name",
                "parent_id":parent_id,
                "parent_name":text,
                "product_code":base["product_code"],
                "name":base["name"],
                "description":base["description"],
                "price":base["price"],
                "delivery":base["delivery"]
            }
            return await bot.send_message(chat_id=cid, text=f"✅ دسته اصلی «{text}» ساخته شد.\n\n📝 حالا نام زیر‌دسته این دسته را وارد کنید:")

        if isinstance(state, dict) and state.get("state") == "product_add_new_root_sub_name":
            if not text:
                return await bot.send_message(chat_id=cid, text="❌ نام زیر‌دسته نمی‌تواند خالی باشد.")
            parent = int(state["parent_id"])
            conn = get_db()
            exists = conn.execute("SELECT id FROM categories WHERE parent_id=? AND name=?", (parent, text)).fetchone()
            parent_row = conn.execute("SELECT id,name,active FROM categories WHERE id=? AND parent_id IS NULL", (parent,)).fetchone()
            if not parent_row or not parent_row["active"]:
                conn.close()
                admin_states.pop(cid, None)
                return await bot.send_message(chat_id=cid, text="❌ دسته اصلی معتبر نیست.", chat_keypad=admin_products_menu(), chat_keypad_type="New")
            if exists:
                conn.close()
                return await bot.send_message(chat_id=cid, text="❌ این زیر‌دسته قبلاً وجود دارد. نام دیگری وارد کنید:")
            cur = conn.execute("INSERT INTO categories(parent_id,name,active,sort_order,created_at) VALUES(?,?,1,0,?)", (parent,text,now_string()))
            cat_id = cur.lastrowid
            conn.commit(); conn.close()
            admin_states[cid] = {
                "state":"product_add", "step":"image", "category_id":cat_id,
                "code":state["product_code"], "name":state["name"],
                "description":state["description"], "price":state["price"],
                "delivery":state["delivery"]
            }
            return await bot.send_message(chat_id=cid, text=f"✅ زیر‌دسته «{text}» زیرِ «{state['parent_name']}» ساخته و برای محصول انتخاب شد.\n\n🖼 حالا تصویر محصول را ارسال کنید. اگر تصویر ندارید، بنویسید: بدون تصویر")

        if isinstance(state, dict) and state.get("state") == "product_add_new_sub_name":
            if not text:
                return await bot.send_message(chat_id=cid, text="❌ نام زیر‌دسته نمی‌تواند خالی باشد.")
            parent = int(state["parent_id"])
            conn = get_db()
            parent_row = conn.execute("SELECT id,name,active FROM categories WHERE id=? AND parent_id IS NULL", (parent,)).fetchone()
            exists = conn.execute("SELECT id FROM categories WHERE parent_id=? AND name=?", (parent, text)).fetchone()
            if not parent_row or not parent_row["active"]:
                conn.close()
                admin_states.pop(cid, None)
                return await bot.send_message(chat_id=cid, text="❌ دسته اصلی معتبر نیست.", chat_keypad=admin_products_menu(), chat_keypad_type="New")
            if exists:
                conn.close()
                return await bot.send_message(chat_id=cid, text="❌ این زیر‌دسته قبلاً وجود دارد. از همان گزینه قبلی استفاده کنید.", chat_keypad=admin_product_category_menu(), chat_keypad_type="New")
            cur = conn.execute("INSERT INTO categories(parent_id,name,active,sort_order,created_at) VALUES(?,?,1,0,?)", (parent,text,now_string()))
            cat_id = cur.lastrowid
            conn.commit(); conn.close()
            base = state
            admin_states[cid] = {"state":"product_add", "step":"image", "category_id":cat_id,
                                 "code":base["product_code"], "name":base["name"], "description":base["description"],
                                 "price":base["price"], "delivery":base["delivery"]}
            return await bot.send_message(chat_id=cid, text=f"✅ زیر‌دسته «{text}» زیرِ «{parent_row['name']}» ساخته و برای محصول انتخاب شد.\n\n🖼 حالا تصویر محصول را ارسال کنید. اگر تصویر ندارید، بنویسید: بدون تصویر")

        if isinstance(state, dict) and state.get("state") == "product_add_new_sub_parent":
            if not norm(text).isdigit():
                return await bot.send_message(chat_id=cid, text="❌ شناسه دسته اصلی باید عدد باشد.", chat_keypad=admin_product_category_menu(), chat_keypad_type="New")
            parent = int(norm(text))
            conn = get_db(); row = conn.execute("SELECT id,name,active FROM categories WHERE id=? AND parent_id IS NULL", (parent,)).fetchone(); conn.close()
            if not row or not row["active"]:
                return await bot.send_message(chat_id=cid, text="❌ دسته اصلی پیدا نشد یا غیرفعال است.", chat_keypad=admin_product_category_menu(), chat_keypad_type="New")
            base = admin_states[cid]
            admin_states[cid] = {"state":"product_add_new_sub_name", "parent_id":parent,
                                 "product_code":base["product_code"], "name":base["name"], "description":base["description"],
                                 "price":base["price"], "delivery":base["delivery"]}
            return await bot.send_message(chat_id=cid, text=f"📝 نام زیر‌دسته جدید برای «{row['name']}» را وارد کنید:")

        if isinstance(state, dict) and state.get("state") == "product_add":
            step = state.get("step")
            if step == "code":
                code = text.upper()
                conn = get_db(); exists = conn.execute("SELECT id FROM products WHERE code=?", (code,)).fetchone(); conn.close()
                if exists:
                    return await bot.send_message(chat_id=cid, text="❌ این کد محصول قبلاً وجود دارد. کد دیگری وارد کنید.")
                state.update(step="name", code=code, product_code=code)
                return await bot.send_message(chat_id=cid, text="📝 نام محصول:")
            if step == "name":
                state.update(step="description", name=text); return await bot.send_message(chat_id=cid, text="📄 توضیحات محصول:")
            if step == "description":
                state.update(step="price", description=text); return await bot.send_message(chat_id=cid, text="💰 قیمت به تومان:")
            if step == "price":
                if not norm(text).isdigit(): return await bot.send_message(chat_id=cid, text="❌ قیمت باید عدد باشد.")
                state.update(step="delivery", price=norm(text)); return await bot.send_message(chat_id=cid, text="⏱ زمان ارسال (مثلاً 1 تا 3):")
            if step == "delivery":
                state.update(step="category", delivery=text)
                return await bot.send_message(
                    chat_id=cid,
                    text="🗂 دسته یا زیر‌دسته محصول را انتخاب کنید.\n\nبرای تعریف دسته جدید، ابتدا به «مدیریت دسته‌بندی‌ها» برگردید.",
                    chat_keypad=admin_product_category_menu(),
                    chat_keypad_type="New"
                )
            if step == "category":
                if not norm(text).isdigit():
                    return await bot.send_message(
                        chat_id=cid,
                        text="❌ لطفاً یکی از دکمه‌های دسته‌بندی را انتخاب کنید یا شناسه عددی معتبر وارد کنید.",
                        chat_keypad=admin_product_category_menu(),
                        chat_keypad_type="New"
                    )
                cat = int(norm(text))
                conn = get_db()
                exists = conn.execute(
                    "SELECT id FROM categories c LEFT JOIN categories p ON p.id=c.parent_id "
                    "WHERE c.id=? AND c.active=1 AND (c.parent_id IS NULL OR p.active=1)",
                    (cat,)
                ).fetchone()
                conn.close()
                if not exists:
                    return await bot.send_message(
                        chat_id=cid,
                        text="❌ دسته پیدا نشد یا غیرفعال است.",
                        chat_keypad=admin_product_category_menu(),
                        chat_keypad_type="New"
                    )
                state.update(step="image", category_id=cat)
                return await bot.send_message(chat_id=cid, text="🖼 تصویر محصول را ارسال کنید. اگر تصویر ندارید، بنویسید: بدون تصویر")
            if step == "image":
                image_path = ""
                if text != "بدون تصویر":
                    raw = getattr(message, "raw_data", None) or {}
                    def find_fid(obj):
                        if isinstance(obj, dict):
                            for key in ("file_id", "fileId"):
                                if obj.get(key): return str(obj[key])
                            for value in obj.values():
                                found = find_fid(value)
                                if found: return found
                        elif isinstance(obj, list):
                            for value in obj:
                                found = find_fid(value)
                                if found: return found
                        return None
                    fid = find_fid(raw)
                    if fid:
                        filename = f"product_{state['code']}.jpg"
                        try:
                            await bot.download(file_id=fid, save_as=filename, verbose=False)
                            image_path = filename
                        except Exception:
                            image_path = ""
                conn = get_db()
                conn.execute("""
                    INSERT INTO products(code,name,description,price,delivery_days,image_path,category_id,active,created_at)
                    VALUES(?,?,?,?,?,?,?,?,?)
                """, (state["code"],state["name"],state["description"],state["price"],state["delivery"],image_path,state["category_id"],1,now_string()))
                conn.commit(); conn.close(); admin_states.pop(cid, None)
                return await bot.send_message(chat_id=cid, text=f"✅ محصول {state['code']} با موفقیت اضافه شد.", chat_keypad=admin_products_menu(), chat_keypad_type="New")

        if isinstance(state, dict) and state.get("state") == "edit_product":
            code = state["code"]; field = state["field"]
            conn = get_db()
            if field == "price":
                if not norm(text).isdigit(): conn.close(); return await bot.send_message(chat_id=cid, text="❌ قیمت باید عدد باشد.")
                value = norm(text); dbfield = "price"
            elif field == "name": value = text; dbfield = "name"
            elif field == "description": value = text; dbfield = "description"
            elif field == "delivery": value = text; dbfield = "delivery_days"
            elif field == "category":
                if not norm(text).isdigit(): conn.close(); return await bot.send_message(chat_id=cid, text="❌ شناسه دسته باید عدد باشد.")
                value = int(norm(text)); dbfield = "category_id"
            else:
                conn.close(); admin_states.pop(cid,None); return
            conn.execute(f"UPDATE products SET {dbfield}=? WHERE code=?", (value,code)); conn.commit(); conn.close(); admin_states.pop(cid,None)
            return await bot.send_message(chat_id=cid, text="✅ محصول به‌روزرسانی شد.", chat_keypad=admin_products_menu(), chat_keypad_type="New")

    # ---------------- CUSTOMER MANUAL ORDER STATE ----------------
    if uid in customers:
        st = customers[uid]
        if st["step"] == "name": st["name"] = text; st["step"] = "phone"; return await bot.send_message(chat_id=cid, text="📱 شماره تماس:")
        if st["step"] == "phone": st["phone"] = norm(text); st["step"] = "address"; return await bot.send_message(chat_id=cid, text="📍 آدرس کامل:")
        if st["step"] == "address":
            st["address"] = text; st["step"] = "confirm"
            return await bot.send_message(chat_id=cid, text=f"🧾 اطلاعات سفارش\n\n👤 {st['name']}\n📱 {st['phone']}\n📍 {st['address']}\n📦 {st['product']}\n\nآیا سفارش تأیید شود؟", chat_keypad=confirm_menu(), chat_keypad_type="New")

    # ---------------- AFTER SALES STATE ----------------
    if uid in after_sales_states:
        st = after_sales_states[uid]
        if text == "لغو":
            after_sales_states.pop(uid,None); return await bot.send_message(chat_id=cid, text="❌ عملیات لغو شد.", chat_keypad=main_menu(), chat_keypad_type="New")
        if st["step"] == "reason":
            st["reason"] = text; oid = st["order_id"]
            conn = get_db(); order = conn.execute("SELECT user_id,product,name,phone,price FROM orders o LEFT JOIN products p ON p.code=o.product WHERE o.id=?", (oid,)).fetchone(); conn.close()
            if not order or order["user_id"] != uid:
                after_sales_states.pop(uid,None); return
            if st["type"] == "تعویض":
                conn=get_db(); cur=conn.execute("INSERT OR IGNORE INTO return_requests(order_id,user_id,product,request_type,reason,status,created_at,return_deadline,customer_name,phone,refund_amount) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(oid,uid,order["product"],"تعویض",text,"در انتظار بررسی",now_string(),deadline(st.get("accepted_at") or now_string(),EXCHANGE_WINDOW_HOURS).strftime("%Y-%m-%d %H:%M:%S"),order["name"],order["phone"],order["price"] or "")); rid=cur.lastrowid; conn.commit(); conn.close(); after_sales_states.pop(uid,None)
                await bot.send_message(chat_id=cid,text=f"✅ درخواست تعویض #{rid} ثبت شد و برای مدیر ارسال شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_after_sales(rid)
            st["step"]="bank"; return await bot.send_message(chat_id=cid,text="🏦 نام صاحب حساب:")
        if st["step"] == "bank": st["account_name"]=text; st["step"]="card"; return await bot.send_message(chat_id=cid,text="💳 شماره کارت ۱۶ رقمی:")
        if st["step"] == "card":
            card=norm(text).replace("-","").replace(" ","")
            if len(card)!=16 or not card.isdigit(): return await bot.send_message(chat_id=cid,text="❌ شماره کارت باید ۱۶ رقم باشد.")
            st["card"]=card; st["step"]="iban"; return await bot.send_message(chat_id=cid,text="🏦 شماره شبا (IR...):")
        if st["step"] == "iban": st["iban"]=norm(text).replace(" ","").upper(); st["step"]="bank_name"; return await bot.send_message(chat_id=cid,text="🏦 نام بانک:")
        if st["step"] == "bank_name":
            conn=get_db(); order=conn.execute("SELECT user_id,product,name,phone FROM orders WHERE id=?",(st["order_id"],)).fetchone(); prod=conn.execute("SELECT price FROM products WHERE code=?",(order["product"],)).fetchone() if order else None
            if not order or order["user_id"] != uid: conn.close(); after_sales_states.pop(uid,None); return
            existing=conn.execute("SELECT id FROM return_requests WHERE order_id=?",(st["order_id"],)).fetchone()
            data=(uid,order["product"],"مرجوعی و بازگشت وجه",st["reason"],"در انتظار بررسی",now_string(),deadline(st.get("accepted_at") or now_string(),RETURN_WINDOW_HOURS).strftime("%Y-%m-%d %H:%M:%S"),order["name"],order["phone"],prod[0] if prod else "",st["account_name"],st["card"],st["iban"],text)
            if existing:
                conn.execute("UPDATE return_requests SET user_id=?,product=?,request_type=?,reason=?,status=?,created_at=?,return_deadline=?,customer_name=?,phone=?,refund_amount=?,bank_account_name=?,card_number=?,iban=?,bank_name=?,decided_at=NULL,admin_note=NULL,paid_at=NULL WHERE id=?", data+(existing[0],))
                rid=existing[0]
            else:
                cur=conn.execute("INSERT INTO return_requests(order_id,user_id,product,request_type,reason,status,created_at,return_deadline,customer_name,phone,refund_amount,bank_account_name,card_number,iban,bank_name) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(st["order_id"],)+data); rid=cur.lastrowid
            conn.commit(); conn.close(); after_sales_states.pop(uid,None)
            await bot.send_message(chat_id=cid,text=f"✅ درخواست مرجوعی #{rid} ثبت شد و برای مدیر ارسال شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_after_sales(rid)

    if rating_comments.get(uid):
        data=rating_comments.pop(uid); conn=get_db(); conn.execute("INSERT INTO reviews(order_id,user_id,product,rating,comment,created_at) VALUES(?,?,?,?,?,?)",(data["order_id"],uid,data["product"],data["rating"],"" if text=="بدون نظر" else text,now_string())); conn.commit(); conn.close(); return await bot.send_message(chat_id=cid,text="🙏 ممنون از ثبت نظر شما.",chat_keypad=main_menu(),chat_keypad_type="New")

# ============================================================
# CALLBACK HANDLER
# ============================================================
@bot.on_callback()
async def handle_callback(bot_instance: Robot, message: Message):
    bid=button_id(message); uid=get_user_id(message); cid=get_chat_id(message)
    if not bid: return

    # ---------------- ADMIN ----------------
    if bid.startswith("admin_"):
        if not is_admin(message): return
        if bid == "admin_back": return await bot.send_message(chat_id=cid,text="🏠 پنل اصلی مدیریت",chat_keypad=admin_menu(),chat_keypad_type="New")
        if bid == "admin_customer_panel": return await bot.send_message(chat_id=cid,text="👤 امور مشتریان و سفارش‌ها",chat_keypad=admin_customer_menu(),chat_keypad_type="New")
        if bid == "admin_categories": return await bot.send_message(chat_id=cid,text="🗂 مدیریت دسته‌بندی‌ها",chat_keypad=admin_categories_menu(),chat_keypad_type="New")
        if bid == "admin_products": return await bot.send_message(chat_id=cid,text="📦 مدیریت محصولات",chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid == "admin_category_add": admin_states[cid]="category_name"; return await bot.send_message(chat_id=cid,text="📝 نام دسته اصلی را وارد کنید:\nبرای لغو: لغو")
        if bid == "admin_subcategory_add": admin_states[cid]="subcategory_parent"; return await bot.send_message(chat_id=cid,text=category_list_text()+"\n\n📝 شناسه دسته اصلی را وارد کنید:")
        if bid == "admin_category_list": return await bot.send_message(chat_id=cid,text=category_list_text(),chat_keypad=admin_categories_menu(),chat_keypad_type="New")
        if bid.startswith("admin_product_add_cat_parent_"):
            parent = int(bid.replace("admin_product_add_cat_parent_", ""))
            state = admin_states.get(cid)
            if not isinstance(state, dict) or state.get("state") != "product_add_new_sub_parent":
                return await bot.send_message(chat_id=cid, text="❌ فرآیند افزودن محصول منقضی شده است.", chat_keypad=admin_products_menu(), chat_keypad_type="New")
            conn=get_db(); row=conn.execute("SELECT id,name,active FROM categories WHERE id=? AND parent_id IS NULL",(parent,)).fetchone(); conn.close()
            if not row or not row["active"]:
                return await bot.send_message(chat_id=cid,text="❌ دسته اصلی معتبر نیست.",chat_keypad=admin_categories_menu(),chat_keypad_type="New")
            state["state"]="product_add_new_sub_name"; state["parent_id"]=parent
            return await bot.send_message(chat_id=cid,text=f"📝 نام زیر‌دسته جدید برای «{row['name']}» را وارد کنید:\nبرای لغو: لغو")
        if bid.startswith("admin_product_add_cat_"):
            cat_id = int(bid.replace("admin_product_add_cat_", ""))
            state = admin_states.get(cid)
            if not isinstance(state, dict) or state.get("state") != "product_add":
                return await bot.send_message(chat_id=cid, text="❌ فرآیند افزودن محصول منقضی شده است.", chat_keypad=admin_products_menu(), chat_keypad_type="New")
            conn = get_db()
            valid = conn.execute("SELECT c.id,c.name,c.parent_id FROM categories c LEFT JOIN categories p ON p.id=c.parent_id WHERE c.id=? AND c.active=1 AND (c.parent_id IS NULL OR p.active=1)", (cat_id,)).fetchone()
            conn.close()
            if not valid:
                return await bot.send_message(chat_id=cid, text="❌ این دسته قابل انتخاب نیست.", chat_keypad=admin_product_category_menu(), chat_keypad_type="New")
            state["step"] = "code"
            state["category_id"] = cat_id
            return await bot.send_message(chat_id=cid, text=f"✅ دسته «{valid['name']}» انتخاب شد.\n\n🔢 کد محصول جدید (مثلاً P-002):")
        if bid == "admin_product_add":
            admin_states[cid]={"state":"product_add","step":"category"}
            return await bot.send_message(
                chat_id=cid,
                text="🗂 ابتدا دسته یا زیر‌دسته محصول را انتخاب کنید.\n\nاگر دسته موردنظر در فهرست نیست، ابتدا از «مدیریت دسته‌بندی‌ها» آن را تعریف کنید و سپس دوباره افزودن محصول را بزنید.",
                chat_keypad=admin_product_category_menu(),
                chat_keypad_type="New"
            )
        if bid == "admin_product_list": return await admin_product_list(cid)
        if bid == "admin_inactive_products": return await admin_inactive_product_list(cid)
        if bid == "admin_edit_product_select": return await admin_product_selection(cid, "edit")
        if bid == "admin_delete_product_select": return await admin_product_selection(cid, "delete")
        if bid.startswith("admin_toggle_product_"):
            code=bid.replace("admin_toggle_product_","")
            conn=get_db(); row=conn.execute("SELECT active FROM products WHERE code=?",(code,)).fetchone()
            if not row:
                conn.close(); return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            new_active=0 if row[0] else 1
            conn.execute("UPDATE products SET active=? WHERE code=?",(new_active,code)); conn.commit(); conn.close()
            msg="✅ محصول فعال شد و به لیست محصولات فعال برگشت." if new_active else "🚫 محصول غیرفعال شد و به لیست محصولات غیرفعال منتقل شد."
            return await bot.send_message(chat_id=cid,text=msg,chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid.startswith("admin_select_edit_"):
            code=bid.replace("admin_select_edit_","")
            conn=get_db(); row=conn.execute("SELECT code,name FROM products WHERE code=?",(code,)).fetchone(); conn.close()
            if not row: return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            # وضعیت محصول را از همان رکوردی که برای انتخاب خوانده‌ایم دریافت می‌کنیم.
            conn = get_db()
            current = conn.execute("SELECT active FROM products WHERE code=?", (code,)).fetchone()
            conn.close()
            if not current:
                return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            active = bool(current[0])
            status_button_id = f"admin_toggle_product_{code}"
            status_button_text = "🚫 غیرفعال کردن محصول" if active else "♻️ فعال‌سازی محصول"
            keypad=(ChatKeypadBuilder()
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_name_{code}",text="📝 نام"),ChatKeypadBuilder().button(id=f"admin_edit_field_price_{code}",text="💰 قیمت"))
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_description_{code}",text="📄 توضیحات"),ChatKeypadBuilder().button(id=f"admin_edit_field_delivery_{code}",text="⏱ ارسال"))
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_category_{code}",text="🗂 دسته"))
                .row(ChatKeypadBuilder().button(id=status_button_id,text=status_button_text))
                .row(ChatKeypadBuilder().button(id="admin_products",text="↩️ محصولات"),ChatKeypadBuilder().button(id="admin_back",text="🏠 خانه"))
                .build())
            status_text = "🟢 فعال" if active else "🔴 غیرفعال"
            return await bot.send_message(chat_id=cid,text=f"✏️ ویرایش محصول {row['code']}\n{row['name']}\nوضعیت: {status_text}\n\nلطفاً مشخصه موردنظر را انتخاب کنید.",chat_keypad=keypad,chat_keypad_type="New")
        if bid.startswith("admin_select_delete_"):
            code=bid.replace("admin_select_delete_","")
            conn=get_db(); row=conn.execute("SELECT code,name,active FROM products WHERE code=?",(code,)).fetchone(); conn.close()
            if not row: return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            keypad=ChatKeypadBuilder().row(ChatKeypadBuilder().button(id=f"admin_confirm_delete_{code}",text="✅ حذف/انتقال به غیرفعال"),ChatKeypadBuilder().button(id="admin_products",text="❌ انصراف")).build()
            return await bot.send_message(chat_id=cid,text=f"⚠️ حذف محصول\n\nکد: {row['code']}\nنام: {row['name']}\n\nمحصول از لیست فعال خارج و در «محصولات غیرفعال» نگهداری می‌شود تا سوابق سفارش‌ها حفظ شود.\n\nآیا مطمئن هستید؟",chat_keypad=keypad,chat_keypad_type="New")
        if bid.startswith("admin_confirm_delete_"):
            code=bid.replace("admin_confirm_delete_","")
            conn=get_db(); row=conn.execute("SELECT code,name FROM products WHERE code=?",(code,)).fetchone()
            if not row:
                conn.close(); return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            conn.execute("UPDATE products SET active=0 WHERE code=?",(code,)); conn.commit(); conn.close()
            return await bot.send_message(chat_id=cid,text=f"🗑 محصول {code} از لیست فعال حذف و به محصولات غیرفعال منتقل شد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
        if bid.startswith("admin_edit_product_"):
            code=bid.replace("admin_edit_product_","")
            conn=get_db()
            current=conn.execute("SELECT active FROM products WHERE code=?",(code,)).fetchone()
            conn.close()
            if not current:
                return await bot.send_message(chat_id=cid,text="❌ محصول پیدا نشد.",chat_keypad=admin_products_menu(),chat_keypad_type="New")
            active=bool(current[0])
            status_button_text="🚫 غیرفعال کردن محصول" if active else "♻️ فعال‌سازی محصول"
            keypad=(ChatKeypadBuilder()
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_name_{code}",text="📝 نام"),ChatKeypadBuilder().button(id=f"admin_edit_field_price_{code}",text="💰 قیمت"))
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_description_{code}",text="📄 توضیحات"),ChatKeypadBuilder().button(id=f"admin_edit_field_delivery_{code}",text="⏱ ارسال"))
                .row(ChatKeypadBuilder().button(id=f"admin_edit_field_category_{code}",text="🗂 دسته"))
                .row(ChatKeypadBuilder().button(id=f"admin_toggle_product_{code}",text=status_button_text))
                .row(ChatKeypadBuilder().button(id="admin_products",text="↩️ محصولات"),ChatKeypadBuilder().button(id="admin_back",text="🏠 خانه"))
                .build())
            return await bot.send_message(chat_id=cid,text=f"✏️ ویرایش {code}\nوضعیت: {'🟢 فعال' if active else '🔴 غیرفعال'}\n\nلطفاً بخش موردنظر را انتخاب کنید.",chat_keypad=keypad,chat_keypad_type="New")
        if bid.startswith("admin_edit_field_"):
            parts=bid.split("_"); field=parts[3]; code="_".join(parts[4:]); admin_states[cid]={"state":"edit_product","field":field,"code":code}
            prompt={"name":"نام جدید:","price":"قیمت جدید:","description":"توضیحات جدید:","delivery":"زمان ارسال جدید:","category":"شناسه دسته/زیر‌دسته جدید:"}.get(field,"مقدار جدید:")
            return await bot.send_message(chat_id=cid,text=prompt+"\n\nبرای لغو: لغو")
        if bid == "admin_export_orders": return await export_excel(cid)
        if bid == "admin_orders":
            conn=get_db(); rows=conn.execute("SELECT id,name,product,status FROM orders ORDER BY id DESC LIMIT 30").fetchall(); conn.close()
            text="📋 آخرین سفارش‌ها\n\n"+"\n".join(f"#{r['id']} | {r['name']} | {r['product']} | {status_fa(r['status'])}" for r in rows) if rows else "📋 سفارشی وجود ندارد."
            return await bot.send_message(chat_id=cid,text=text,chat_keypad=admin_customer_menu(),chat_keypad_type="New")
        if bid == "admin_after_sales":
            conn=get_db(); rows=conn.execute("SELECT id,order_id,request_type,status FROM return_requests WHERE status NOT IN ('تکمیل شده','تعویض انجام شد','وجه واریز شد','رد شده') ORDER BY id DESC").fetchall(); conn.close()
            if not rows: return await bot.send_message(chat_id=cid,text="✅ درخواست باز خدمات پس از فروش وجود ندارد.",chat_keypad=admin_customer_menu(),chat_keypad_type="New")
            for r in rows: await bot.send_message(chat_id=cid,text=f"🔄 درخواست #{r['id']}\n🧾 سفارش #{r['order_id']}\nنوع: {r['request_type']}\nوضعیت: {after_status(r['status'])}",chat_keypad=return_admin_menu(r['id'],r['request_type'],r['status']),chat_keypad_type="New")
            return
        if bid.startswith("admin_confirm_"):
            oid=int(bid.split("_")[-1]); conn=get_db(); conn.execute("UPDATE orders SET status='تأیید شده' WHERE id=?",(oid,)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"✅ سفارش #{oid} تأیید شد."); return await notify_customer_status(oid,"تأیید شده")
        if bid.startswith("admin_cancel_"):
            oid=int(bid.split("_")[-1]); conn=get_db(); conn.execute("UPDATE orders SET status='لغو شده' WHERE id=?",(oid,)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"❌ سفارش #{oid} لغو شد."); return await notify_customer_status(oid,"لغو شده")
        if bid.startswith("admin_deliver_"):
            oid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT chat_id,name FROM orders WHERE id=?",(oid,)).fetchone()
            if not row: conn.close(); return
            conn.execute("UPDATE orders SET status='تحویل شده',delivered_at=? WHERE id=?",(now_string(),oid)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"🚚 سفارش #{oid} تحویل شد."); return await bot.send_message(chat_id=row['chat_id'],text=f"سلام {row['name']} عزیز 🌷\n\n📦 سفارش #{oid} تحویل شما شده است. پس از رویت و بررسی، دریافت کالا را تأیید کنید.",chat_keypad=accept_menu(oid),chat_keypad_type="New")
        if bid.startswith("admin_after_approve_") or bid.startswith("admin_after_reject_"):
            rid=int(bid.split("_")[-1]); approve="approve" in bid; conn=get_db(); row=conn.execute("SELECT order_id,request_type FROM return_requests WHERE id=?",(rid,)).fetchone()
            if not row: conn.close(); return
            if approve: new_status="در انتظار واریز" if row['request_type']=="مرجوعی و بازگشت وجه" else "در حال تعویض"
            else: new_status="رد شده"
            note="درخواست توسط مدیر تأیید شد." if approve else "درخواست توسط مدیر رد شد."
            conn.execute("UPDATE return_requests SET status=?,decided_at=?,admin_note=?,exchange_delivery_days=? WHERE id=?",(new_status,now_string(),note,"2 تا 3" if row['request_type']=="تعویض" and approve else None,rid)); conn.commit(); conn.close()
            await bot.send_message(chat_id=cid,text=f"{'✅' if approve else '❌'} درخواست #{rid} بررسی شد.")
            await notify_customer_after(rid); return
        if bid.startswith("admin_refund_paid_"):
            rid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT status FROM return_requests WHERE id=? AND request_type='مرجوعی و بازگشت وجه'",(rid,)).fetchone()
            if not row or row['status']!="در انتظار واریز": conn.close(); return
            conn.execute("UPDATE return_requests SET status='وجه واریز شد',paid_at=?,decided_at=? WHERE id=?",(now_string(),now_string(),rid)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"✅ واریز درخواست #{rid} ثبت شد."); return await notify_customer_after(rid)
        if bid.startswith("admin_exchange_sent_"):
            rid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT status FROM return_requests WHERE id=? AND request_type='تعویض'",(rid,)).fetchone()
            if not row or row['status']!="در حال تعویض": conn.close(); return
            conn.execute("UPDATE return_requests SET status='کالای تعویضی ارسال شد',exchange_sent_at=?,decided_at=? WHERE id=?",(now_string(),now_string(),rid)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"📦 ارسال کالای تعویضی درخواست #{rid} ثبت شد."); return await notify_customer_after(rid)
        return

    # ---------------- CUSTOMER ----------------
    if bid == "main_menu":
        customers.pop(uid,None); after_sales_states.pop(uid,None); return await bot.send_message(chat_id=cid,text="🏪 سبلان شاپ\n\nلطفاً گزینه موردنظر را انتخاب کنید.",chat_keypad=main_menu(),chat_keypad_type="New")
    if bid == "show_categories": return await bot.send_message(chat_id=cid,text="🗂 دسته‌بندی محصولات",chat_keypad=category_menu(),chat_keypad_type="New")
    if bid.startswith("cat_"):
        cat_id=int(bid.split("_")[-1]); keypad, rows=products_in_category_menu(cat_id)
        conn=get_db(); cat=conn.execute("SELECT name FROM categories WHERE id=?",(cat_id,)).fetchone(); conn.close()
        if rows:
            return await bot.send_message(chat_id=cid,text=f"📂 {cat['name']}\n\nمحصول موردنظر را انتخاب کنید:",chat_keypad=keypad,chat_keypad_type="New")
        # Show children when no direct products.
        return await bot.send_message(chat_id=cid,text=f"📂 {cat['name']}\n\nزیر‌دسته یا محصول را انتخاب کنید:",chat_keypad=category_menu(cat_id),chat_keypad_type="New")
    if bid.startswith("view_product_"):
        return await show_product(cid,bid.replace("view_product_",""))
    if bid.startswith("order_product_"):
        code=bid.replace("order_product_","")
        if not get_product(code): return await bot.send_message(chat_id=cid,text="❌ محصول موجود نیست.",chat_keypad=main_menu(),chat_keypad_type="New")
        link=website_order_link(cid,code)
        return await bot.send_message(
            chat_id=cid,
            text=(
                f"🛒 ادامه فرآیند خرید\n\n"
                f"محصول انتخابی: {code}\n\n"
                "⚠️ به دلیل استفاده از زیرساخت بین‌المللی سایت فروشگاه، "
                "ممکن است برای باز شدن صفحه فروشگاه نیاز باشد اتصال اینترنت "
                "خود را از طریق VPN فعال کنید.\n\n"
                f"🔗 لینک ادامه خرید:\n{link}\n\n"
                "شناسه مشتری شما به‌صورت خودکار همراه سفارش منتقل می‌شود."
            )
        )
    if bid == "shop_policy": return await bot.send_message(chat_id=cid,text=policy_text(),chat_keypad=main_menu(),chat_keypad_type="New")
    if bid == "my_orders":
        conn=get_db(); rows=conn.execute("SELECT id,product,status,created_at,accepted_at FROM orders WHERE user_id=? OR chat_id=? ORDER BY id DESC",(uid,cid)).fetchall(); conn.close()
        if not rows: return await bot.send_message(chat_id=cid,text="📋 هنوز سفارشی برای شما ثبت نشده است.",chat_keypad=main_menu(),chat_keypad_type="New")
        for row in rows:
            text=f"🧾 سفارش #{row['id']}\n📦 {row['product']}\n📌 {status_fa(row['status'])}\n📅 {row['created_at'] or '-'}"
            keypad=None
            if row['status']=="تحویل و تأیید شده": keypad=customer_after_menu(row['id'])
            await bot.send_message(chat_id=cid,text=text,chat_keypad=keypad,chat_keypad_type="New" if keypad else None)
        return
    if bid == "customer_confirm_order":
        st=customers.get(uid)
        if not st or st.get("step")!="confirm": return
        conn=get_db(); cur=conn.execute("INSERT INTO orders(user_id,product,name,phone,address,status,chat_id,created_at) VALUES(?,?,?,?,?,?,?,?)",(uid,st['product'],st['name'],st['phone'],st['address'],"جدید",cid,now_string())); oid=cur.lastrowid; conn.commit(); conn.close(); customers.pop(uid,None); await bot.send_message(chat_id=cid,text=f"✅ سفارش #{oid} ثبت شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await notify_admin_order(oid)
    if bid == "customer_cancel_order": customers.pop(uid,None); return await bot.send_message(chat_id=cid,text="❌ سفارش لغو شد.",chat_keypad=main_menu(),chat_keypad_type="New")
    if bid.startswith("customer_accept_delivery_"):
        oid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT user_id,status FROM orders WHERE id=?",(oid,)).fetchone()
        if not row or row['user_id']!=uid or row['status']!="تحویل شده": conn.close(); return
        accepted=now_string(); conn.execute("UPDATE orders SET status='تحویل و تأیید شده',accepted_at=?,customer_acceptance=? WHERE id=?",(accepted,"مشتری کالا را رویت و سالم و مطابق سفارش تحویل گرفت.",oid)); conn.commit(); conn.close(); d=deadline(accepted); await bot.send_message(chat_id=cid,text=f"✅ دریافت سفارش #{oid} ثبت شد.\n\n⏰ مهلت خدمات پس از فروش عادی تا: {d.strftime('%Y-%m-%d %H:%M:%S')}",chat_keypad=customer_after_menu(oid),chat_keypad_type="New"); return await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"📦 مشتری دریافت سفارش #{oid} را تأیید کرد.")
    if bid.startswith("customer_rate_"):
        oid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT user_id,status,product FROM orders WHERE id=?",(oid,)).fetchone(); already=conn.execute("SELECT id FROM reviews WHERE order_id=?",(oid,)).fetchone(); conn.close()
        if not row or row['user_id']!=uid or row['status'] not in ("تحویل و تأیید شده","تعویض انجام شد") or already: return
        return await bot.send_message(chat_id=cid,text=f"⭐ میزان رضایت از سفارش #{oid} را انتخاب کنید:",chat_keypad=rating_menu(oid),chat_keypad_type="New")
    if bid.startswith("rating_"):
        parts=bid.split("_"); oid=int(parts[1]); rating=int(parts[2]); conn=get_db(); row=conn.execute("SELECT user_id,product,status FROM orders WHERE id=?",(oid,)).fetchone(); exists=conn.execute("SELECT id FROM reviews WHERE order_id=?",(oid,)).fetchone(); conn.close()
        if not row or row['user_id']!=uid or row['status'] not in ("تحویل و تأیید شده","تعویض انجام شد") or exists: return
        rating_comments[uid]={"order_id":oid,"product":row['product'],"rating":rating}; return await bot.send_message(chat_id=cid,text=f"⭐ امتیاز {rating} از 5 ثبت شد.\n\n📝 نظر خود را بنویسید یا بنویسید: بدون نظر")
    if bid.startswith("customer_accept_refund_"):
        rid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT order_id,user_id,status FROM return_requests WHERE id=? AND request_type='مرجوعی و بازگشت وجه'",(rid,)).fetchone()
        if not row or row['user_id']!=uid or row['status']!="وجه واریز شد": conn.close(); return
        conn.execute("UPDATE return_requests SET status='تکمیل شده',decided_at=?,admin_note=? WHERE id=?",(now_string(),"مشتری دریافت وجه را تأیید کرد.",rid)); conn.execute("UPDATE orders SET status='مرجوع و وجه بازگشت داده شد' WHERE id=?",(row['order_id'],)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"✅ دریافت وجه درخواست #{rid} تأیید شد.",chat_keypad=main_menu(),chat_keypad_type="New"); return await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"💰 مشتری دریافت وجه درخواست #{rid} را تأیید کرد.")
    if bid.startswith("customer_accept_exchange_"):
        rid=int(bid.split("_")[-1]); conn=get_db(); row=conn.execute("SELECT order_id,user_id,status FROM return_requests WHERE id=? AND request_type='تعویض'",(rid,)).fetchone()
        if not row or row['user_id']!=uid or row['status']!="کالای تعویضی ارسال شد": conn.close(); return
        conn.execute("UPDATE return_requests SET status='تعویض انجام شد',exchange_received_at=?,exchange_customer_acceptance=?,decided_at=? WHERE id=?",(now_string(),"مشتری کالای تعویضی را دریافت و تأیید کرد.",now_string(),rid)); conn.execute("UPDATE orders SET status='تعویض انجام شد' WHERE id=?",(row['order_id'],)); conn.commit(); conn.close(); await bot.send_message(chat_id=cid,text=f"✅ دریافت کالای تعویضی درخواست #{rid} ثبت شد.",chat_keypad=customer_after_menu(row['order_id'],exchange=True),chat_keypad_type="New"); return await bot.send_message(chat_id=ADMIN_CHAT_ID,text=f"📦 مشتری دریافت کالای تعویضی #{rid} را تأیید کرد.")
    if bid.startswith("customer_exchange_") or bid.startswith("customer_refund_"):
        oid=int(bid.split("_")[-1]); typ="تعویض" if bid.startswith("customer_exchange_") else "مرجوعی و بازگشت وجه"; conn=get_db(); row=conn.execute("SELECT user_id,status,accepted_at FROM orders WHERE id=?",(oid,)).fetchone(); existing=conn.execute("SELECT id,status FROM return_requests WHERE order_id=?",(oid,)).fetchone(); conn.close()
        if not row or row['user_id']!=uid or row['status'] not in ("تحویل و تأیید شده","تعویض انجام شد"): return
        if row['status']=="تحویل و تأیید شده":
            d=deadline(row['accepted_at'],EXCHANGE_WINDOW_HOURS if typ=="تعویض" else RETURN_WINDOW_HOURS)
            if not d or datetime.now()>d: return await bot.send_message(chat_id=cid,text="⏰ مهلت ثبت این درخواست به پایان رسیده است.")
        after_sales_states[uid]={"step":"reason","order_id":oid,"type":typ,"accepted_at":row['accepted_at']}; return await bot.send_message(chat_id=cid,text=f"🔄 درخواست {typ}\n\nدلیل درخواست را کامل بنویسید:\n\nبرای لغو: لغو")

# ============================================================
# START
# ============================================================
if __name__ == "__main__":
    init_db()
    print("BOT STARTING...")
    print("Starting website order synchronization...")
    threading.Thread(target=website_sync_worker, daemon=True).start()
    bot.run()
