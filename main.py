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


# =========================================================
# تنظیمات اصلی
# =========================================================

TOKEN = ""

ADMIN_CHAT_ID = "b0BC4FX0BHkJ080af57b490e17163d49"

DB_PATH = "shop.db"

PRODUCT_CODE = "P-001"

PRODUCT_IMAGE_PATH = "product_p001.jpg"

FREE_DELIVERY_CITY = "تبریز"

RECENT_ORDER_DAYS = 7

RECENT_ORDER_THRESHOLD = 3

# مهلت تعویض / مرجوعی پس از تأیید دریافت کالا
RETURN_WINDOW_HOURS = 24


# =========================================================
# اطلاعات اولیه محصول
# =========================================================

DEFAULT_PRODUCT = {
    "code": "P-001",
    "name": "محصول نمونه",
    "description": "توضیحات محصول را اینجا وارد کنید.",
    "price": "1500000",
    "delivery_days": "1 تا 3",
    "image_path": PRODUCT_IMAGE_PATH,
    "active": 1
}


# =========================================================
# دیتابیس
# =========================================================

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()


cursor.execute("""
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT,
    product TEXT,
    name TEXT,
    phone TEXT,
    address TEXT,
    status TEXT
)
""")

conn.commit()


def get_columns(table_name):
    cursor.execute(f"PRAGMA table_info({table_name})")
    return [row[1] for row in cursor.fetchall()]


def add_column_if_missing(
    table_name,
    column_name,
    column_definition
):
    columns = get_columns(table_name)

    if column_name not in columns:
        cursor.execute(
            f"""
            ALTER TABLE {table_name}
            ADD COLUMN {column_name} {column_definition}
            """
        )
        conn.commit()


# =========================================================
# ستون‌های تکمیلی سفارش‌ها
# =========================================================

add_column_if_missing(
    "orders",
    "chat_id",
    "TEXT"
)

add_column_if_missing(
    "orders",
    "created_at",
    "TEXT"
)

add_column_if_missing(
    "orders",
    "delivered_at",
    "TEXT"
)

add_column_if_missing(
    "orders",
    "accepted_at",
    "TEXT"
)

add_column_if_missing(
    "orders",
    "customer_acceptance",
    "TEXT"
)


# =========================================================
# جدول محصولات
# =========================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE,
    name TEXT,
    description TEXT,
    price TEXT,
    delivery_days TEXT,
    image_path TEXT,
    active INTEGER DEFAULT 1
)
""")

conn.commit()


add_column_if_missing(
    "products",
    "image_file_id",
    "TEXT"
)


# =========================================================
# جدول نظرات
# =========================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS reviews (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER,
    user_id TEXT,
    product TEXT,
    rating INTEGER,
    comment TEXT,
    created_at TEXT
)
""")

conn.commit()


# =========================================================
# جدول درخواست‌های تعویض / مرجوعی
# =========================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS return_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id INTEGER UNIQUE,
    user_id TEXT,
    product TEXT,
    request_type TEXT,
    reason TEXT,
    status TEXT,
    created_at TEXT,
    decided_at TEXT,
    admin_note TEXT
)
""")

conn.commit()


# =========================================================
# ایجاد محصول پیش‌فرض
# =========================================================

cursor.execute(
    "SELECT id FROM products WHERE code = ?",
    (DEFAULT_PRODUCT["code"],)
)

product_exists = cursor.fetchone()

if not product_exists:
    cursor.execute("""
        INSERT INTO products
        (
            code,
            name,
            description,
            price,
            delivery_days,
            image_path,
            active
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        DEFAULT_PRODUCT["code"],
        DEFAULT_PRODUCT["name"],
        DEFAULT_PRODUCT["description"],
        DEFAULT_PRODUCT["price"],
        DEFAULT_PRODUCT["delivery_days"],
        DEFAULT_PRODUCT["image_path"],
        DEFAULT_PRODUCT["active"]
    ))

    conn.commit()


# =========================================================
# ربات
# =========================================================

bot = Robot(token=TOKEN)


# =========================================================
# وضعیت موقت
# =========================================================

customers = {}

rating_comments = {}

admin_states = {}

return_request_states = {}


# =========================================================
# توابع کمکی
# =========================================================

def now_string():
    return datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


def is_admin_message(message):
    return (
        getattr(message, "chat_id", None)
        == ADMIN_CHAT_ID
    )


def get_user_id(message):
    return getattr(
        message,
        "sender_id",
        None
    )


def get_chat_id(message):
    return getattr(
        message,
        "chat_id",
        None
    )


def get_button_id(message):
    try:
        return message.aux_data.button_id
    except Exception:
        return None


def normalize_digits(value):

    if value is None:
        return ""

    translation_table = str.maketrans(
        "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
        "01234567890123456789"
    )

    return value.translate(
        translation_table
    )


def format_price(price):

    if price is None:
        return "نامشخص"

    price_text = str(price).strip()

    normalized = normalize_digits(
        price_text
    )

    try:
        number = int(normalized)

        return f"{number:,}"

    except ValueError:
        return price_text


# =========================================================
# تبدیل زمان
# =========================================================

def parse_datetime(value):

    if not value:
        return None

    try:
        return datetime.strptime(
            value,
            "%Y-%m-%d %H:%M:%S"
        )
    except Exception:
        return None


def get_return_deadline(accepted_at):

    accepted_datetime = parse_datetime(
        accepted_at
    )

    if not accepted_datetime:
        return None

    return (
        accepted_datetime
        + timedelta(
            hours=RETURN_WINDOW_HOURS
        )
    )


def format_datetime_fa(value):

    if not value:
        return "نامشخص"

    return value


def is_return_window_open(accepted_at):

    deadline = get_return_deadline(
        accepted_at
    )

    if not deadline:
        return False

    return datetime.now() <= deadline


# =========================================================
# دریافت محصول فعال
# =========================================================

def get_product():

    cursor.execute("""
        SELECT
            code,
            name,
            description,
            price,
            delivery_days,
            image_path,
            image_file_id
        FROM products
        WHERE code = ?
        AND active = 1
    """, (PRODUCT_CODE,))

    return cursor.fetchone()


# =========================================================
# دریافت محصول برای مدیر
# =========================================================

def get_product_admin():

    cursor.execute("""
        SELECT
            code,
            name,
            description,
            price,
            delivery_days,
            image_path,
            image_file_id,
            active
        FROM products
        WHERE code = ?
    """, (PRODUCT_CODE,))

    return cursor.fetchone()


# =========================================================
# آمار سفارش
# =========================================================

def get_total_orders(product_code):

    cursor.execute("""
        SELECT COUNT(*)
        FROM orders
        WHERE product = ?
    """, (product_code,))

    result = cursor.fetchone()

    return result[0] if result else 0


def get_recent_orders(product_code):

    date_limit = (
        datetime.now()
        - timedelta(
            days=RECENT_ORDER_DAYS
        )
    )

    cursor.execute("""
        SELECT COUNT(*)
        FROM orders
        WHERE product = ?
        AND created_at IS NOT NULL
        AND created_at >= ?
    """, (
        product_code,
        date_limit.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    result = cursor.fetchone()

    return result[0] if result else 0


def get_rating_info(product_code):

    cursor.execute("""
        SELECT AVG(rating), COUNT(*)
        FROM reviews
        WHERE product = ?
    """, (product_code,))

    result = cursor.fetchone()

    if not result:
        return 0, 0

    average = result[0]

    count = result[1]

    if average is None:
        average = 0

    return round(average, 1), count


def status_fa(status):

    mapping = {
        "جدید": "🆕 جدید",
        "تأیید شده": "✅ تأیید شده",
        "لغو شده": "❌ لغو شده",
        "تحویل شده": "🚚 تحویل شده",
        "تحویل و تأیید شده": "📦 تحویل و تأیید شده"
    }

    return mapping.get(
        status,
        status
    )


def return_status_fa(status):

    mapping = {
        "در انتظار بررسی": "🟡 در انتظار بررسی",
        "تأیید شده": "✅ تأیید شده",
        "رد شده": "❌ رد شده"
    }

    return mapping.get(
        status,
        status or "نامشخص"
    )


# =========================================================
# منوی اصلی مشتری
# =========================================================

def main_menu():

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id="show_products",
                text="📦 مشاهده محصولات"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="my_orders",
                text="📋 مشاهده سفارش‌ها"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="shop_policy",
                text="📜 شرایط تعویض و مرجوعی"
            )
        )

        .build()
    )


# =========================================================
# منوی محصول
# =========================================================

def product_menu():

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id="order_product_P-001",
                text="🛒 ثبت سفارش"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="main_menu",
                text="🔙 بازگشت به منوی اصلی"
            )
        )

        .build()
    )


# =========================================================
# تأیید سفارش
# =========================================================

def confirm_menu():

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id="customer_confirm_order",
                text="✅ تأیید سفارش"
            ),

            ChatKeypadBuilder().button(
                id="customer_cancel_order",
                text="❌ لغو سفارش"
            )
        )

        .build()
    )


# =========================================================
# سفارش مدیر
# =========================================================

def admin_order_menu(order_id):

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id=f"admin_confirm_{order_id}",
                text="✅ تأیید سفارش"
            ),

            ChatKeypadBuilder().button(
                id=f"admin_cancel_{order_id}",
                text="❌ لغو سفارش"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id=f"admin_deliver_{order_id}",
                text="🚚 تحویل شد"
            )
        )

        .build()
    )


# =========================================================
# منوی تصمیم‌گیری درخواست مرجوعی
# =========================================================

def admin_return_menu(request_id):

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id=f"admin_return_approve_{request_id}",
                text="✅ تأیید درخواست"
            ),

            ChatKeypadBuilder().button(
                id=f"admin_return_reject_{request_id}",
                text="❌ رد درخواست"
            )
        )

        .build()
    )


# =========================================================
# منوی درخواست تعویض / مرجوعی
# =========================================================

def customer_return_menu(order_id):

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id=f"customer_return_{order_id}",
                text="🔄 درخواست تعویض / مرجوعی"
            )
        )

        .build()
    )


# =========================================================
# منوی مدیریت محصولات
# =========================================================

def admin_products_menu():

    product = get_product_admin()

    active = 0

    if product:
        active = product[7]

    if active:
        toggle_text = (
            "🔴 غیرفعال کردن P-001"
        )
    else:
        toggle_text = (
            "🟢 فعال کردن P-001"
        )

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id="admin_price_P-001",
                text="💰 تغییر قیمت P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_name_P-001",
                text="📝 تغییر نام P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_description_P-001",
                text="📄 تغییر توضیحات P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_image_P-001",
                text="🖼 تغییر تصویر P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_delivery_P-001",
                text="⏱ تغییر زمان ارسال P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_toggle_P-001",
                text=toggle_text
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_product_info_P-001",
                text="📋 مشخصات کامل P-001"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_back",
                text="🔙 بازگشت به پنل مدیریت"
            )
        )

        .build()
    )


# =========================================================
# پنل مدیر
# =========================================================

def admin_menu():

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id="admin_products",
                text="📦 مدیریت محصولات"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_export_orders",
                text="📊 خروجی Excel سفارش‌ها"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id="admin_orders",
                text="📋 مشاهده سفارش‌ها"
            )
        )

        .build()
    )


# =========================================================
# امتیازدهی
# =========================================================

def rating_menu(order_id):

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id=f"rating_{order_id}_1",
                text="⭐ 1"
            ),

            ChatKeypadBuilder().button(
                id=f"rating_{order_id}_2",
                text="⭐ 2"
            ),

            ChatKeypadBuilder().button(
                id=f"rating_{order_id}_3",
                text="⭐ 3"
            )
        )

        .row(
            ChatKeypadBuilder().button(
                id=f"rating_{order_id}_4",
                text="⭐ 4"
            ),

            ChatKeypadBuilder().button(
                id=f"rating_{order_id}_5",
                text="⭐ 5"
            )
        )

        .build()
    )


# =========================================================
# تأیید دریافت کالا
# =========================================================

def customer_accept_menu(order_id):

    return (
        ChatKeypadBuilder()

        .row(
            ChatKeypadBuilder().button(
                id=f"customer_accept_delivery_{order_id}",
                text=(
                    "✅ کالا را رویت کردم و سالم و مطابق "
                    "سفارش تحویل گرفتم"
                )
            )
        )

        .build()
    )


# =========================================================
# متن محصول
# =========================================================

def build_product_text():

    product = get_product()

    if not product:
        return "❌ محصول موردنظر پیدا نشد."

    (
        code,
        name,
        description,
        price,
        delivery_days,
        image_path,
        image_file_id
    ) = product

    total_orders = get_total_orders(code)

    recent_orders = get_recent_orders(code)

    average_rating, review_count = get_rating_info(
        code
    )

    displayed_price = format_price(
        price
    )

    text = (
        "🏪 فروشگاه آنلاین شاپ\n"
        "خریدی آسان و مطمئن برای کلیه همشهریان و هموطنان\n\n"

        f"📦 {code}\n"

        f"🔹 نام محصول: {name}\n\n"

        f"📝 توضیحات محصول:\n"
        f"{description}\n\n"

        f"💰 قیمت: {displayed_price} تومان\n\n"

        f"🛒 تاکنون {total_orders} سفارش برای این محصول ثبت شده است.\n"
    )

    if recent_orders >= RECENT_ORDER_THRESHOLD:
        text += (
            "🔥 این محصول اخیراً مورد توجه مشتریان بوده است.\n"
        )

    text += (
        "\n"

        f"🚚 ارسال رایگان فقط در محدوده شهر "
        f"{FREE_DELIVERY_CITY}\n"

        f"⏱ زمان تقریبی ارسال: "
        f"{delivery_days} روز کاری\n\n"

        "💵 پرداخت درب منزل پس از رویت و تحویل کالا\n\n"

        f"⭐ امتیاز مشتریان: {average_rating} / 5\n"

        f"💬 تعداد نظر ثبت شده: {review_count}\n\n"

        "برای ثبت سفارش روی دکمه زیر بزنید."
    )

    return text


# =========================================================
# نمایش محصول
# =========================================================

async def show_product(chat_id):

    product = get_product()

    if not product:

        await bot.send_message(
            chat_id=chat_id,
            text="❌ محصولی برای نمایش وجود ندارد."
        )

        return

    (
        code,
        name,
        description,
        price,
        delivery_days,
        image_path,
        image_file_id
    ) = product

    product_text = build_product_text()

    # تصویر محلی
    if image_path and os.path.exists(
        image_path
    ):

        try:

            await bot.send_image(
                chat_id=chat_id,
                path=image_path,
                text=product_text,
                chat_keypad=product_menu(),
                chat_keypad_type="New"
            )

            return

        except Exception as error:

            print(
                "LOCAL IMAGE SEND ERROR:",
                error
            )

    # fallback روبیکا
    if image_file_id:

        try:

            await bot.send_image(
                chat_id=chat_id,
                file_id=image_file_id,
                text=product_text,
                chat_keypad=product_menu(),
                chat_keypad_type="New"
            )

            return

        except Exception as error:

            print(
                "RUBIKA IMAGE SEND ERROR:",
                error
            )

    await bot.send_message(
        chat_id=chat_id,
        text=(
            "⚠️ تصویر محصول هنوز ثبت نشده است.\n\n"
            + product_text
        ),
        chat_keypad=product_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# قوانین فروشگاه
# =========================================================

def policy_text():

    return (
        "📜 شرایط تعویض و مرجوعی\n\n"

        "1️⃣ هنگام تحویل، مشتری می‌تواند کالا را رویت و بررسی کند.\n\n"

        "2️⃣ پس از اطمینان از سالم بودن کالا و مطابقت آن "
        "با سفارش، مشتری دریافت کالا را در ربات تأیید می‌کند.\n\n"

        f"3️⃣ از زمان تأیید دریافت کالا، مشتری "
        f"{RETURN_WINDOW_HOURS} ساعت فرصت دارد درخواست "
        "تعویض یا مرجوعی خود را در ربات ثبت کند.\n\n"

        "4️⃣ درخواست تعویض یا مرجوعی برای بررسی به مدیر فروشگاه "
        "ارسال می‌شود و تصمیم نهایی توسط فروشگاه ثبت خواهد شد.\n\n"

        "5️⃣ در صورت وجود ایراد، مغایرت با سفارش یا آسیب‌دیدگی، "
        "موضوع باید در اولین فرصت به فروشگاه اطلاع داده شود.\n\n"

        "6️⃣ پس از پایان مهلت تعیین‌شده، دکمه درخواست همچنان "
        "قابل مشاهده است، اما امکان ثبت درخواست جدید وجود ندارد.\n\n"

        "7️⃣ شرایط تعویض یا مرجوعی بر اساس نوع کالا و وضعیت آن "
        "بررسی می‌شود.\n\n"

        "8️⃣ کالاهایی که به دلیل شرایط خاص، استفاده، آسیب ایجادشده "
        "توسط مشتری یا موارد مشابه امکان تعویض ندارند، باید هنگام "
        "فروش به مشتری اعلام شوند.\n\n"

        "⚠️ این بخش، سیاست فروشگاه است و برای قرارداد یا تعهد "
        "حقوقی خاص، در صورت نیاز باید متن آن توسط مشاور حقوقی بررسی شود."
    )


# =========================================================
# سفارش‌های مشتری
# =========================================================

async def show_my_orders(message):

    user_id = get_user_id(message)

    chat_id = get_chat_id(message)

    cursor.execute("""
        SELECT
            id,
            product,
            status,
            created_at,
            accepted_at
        FROM orders
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT 20
    """, (user_id,))

    orders = cursor.fetchall()

    if not orders:

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "📋 سفارش‌های من\n\n"
                "هنوز سفارشی برای شما ثبت نشده است."
            ),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        return

    text = "📋 سفارش‌های من\n\n"

    for order in orders:

        (
            order_id,
            product,
            status,
            created_at,
            accepted_at
        ) = order

        date_text = (
            created_at
            if created_at
            else "تاریخ ثبت نامشخص"
        )

        text += (
            f"🧾 سفارش #{order_id}\n"
            f"📦 محصول: {product}\n"
            f"📌 وضعیت: {status_fa(status)}\n"
            f"🕐 تاریخ: {date_text}\n"
        )

        if accepted_at:

            deadline = get_return_deadline(
                accepted_at
            )

            if deadline:

                text += (
                    f"🔄 مهلت تعویض/مرجوعی تا: "
                    f"{deadline.strftime('%Y-%m-%d %H:%M:%S')}\n"
                )

        # وضعیت درخواست مرجوعی
        cursor.execute("""
            SELECT
                status,
                request_type,
                created_at
            FROM return_requests
            WHERE order_id = ?
        """, (order_id,))

        return_request = cursor.fetchone()

        if return_request:

            (
                return_status,
                request_type,
                return_created_at
            ) = return_request

            text += (
                f"🔄 درخواست: "
                f"{return_status_fa(return_status)}\n"
                f"📌 نوع درخواست: {request_type}\n"
            )

        text += (
            "──────────────\n"
        )

    await bot.send_message(
        chat_id=chat_id,
        text=text,
        chat_keypad=main_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# سفارش‌های مدیر
# =========================================================

async def show_admin_orders(chat_id):

    cursor.execute("""
        SELECT
            id,
            product,
            name,
            phone,
            address,
            status,
            created_at
        FROM orders
        ORDER BY id DESC
        LIMIT 50
    """)

    orders = cursor.fetchall()

    if not orders:

        await bot.send_message(
            chat_id=chat_id,
            text="📋 هنوز هیچ سفارشی ثبت نشده است.",
            chat_keypad=admin_menu(),
            chat_keypad_type="New"
        )

        return

    text = "📋 آخرین سفارش‌ها\n\n"

    for order in orders:

        (
            order_id,
            product,
            name,
            phone,
            address,
            status,
            created_at
        ) = order

        text += (
            f"🧾 سفارش #{order_id}\n"
            f"📦 محصول: {product}\n"
            f"👤 نام: {name}\n"
            f"📱 تلفن: {phone}\n"
            f"📌 وضعیت: {status_fa(status)}\n"
            f"🕐 تاریخ: {created_at or 'نامشخص'}\n"
        )

        cursor.execute("""
            SELECT
                status,
                request_type,
                created_at
            FROM return_requests
            WHERE order_id = ?
        """, (order_id,))

        return_request = cursor.fetchone()

        if return_request:

            (
                return_status,
                request_type,
                return_created_at
            ) = return_request

            text += (
                f"🔄 درخواست تعویض/مرجوعی: "
                f"{return_status_fa(return_status)}\n"
                f"📌 نوع: {request_type}\n"
            )

        text += (
            "──────────────\n"
        )

    await bot.send_message(
        chat_id=chat_id,
        text=text,
        chat_keypad=admin_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# مشخصات کامل محصول
# =========================================================

async def show_admin_product_info(chat_id):

    product = get_product_admin()

    if not product:

        await bot.send_message(
            chat_id=chat_id,
            text="❌ محصول P-001 پیدا نشد.",
            chat_keypad=admin_products_menu(),
            chat_keypad_type="New"
        )

        return

    (
        code,
        name,
        description,
        price,
        delivery_days,
        image_path,
        image_file_id,
        active
    ) = product

    total_orders = get_total_orders(code)

    recent_orders = get_recent_orders(code)

    average_rating, review_count = get_rating_info(
        code
    )

    status = (
        "🟢 فعال"
        if active
        else
        "🔴 غیرفعال"
    )

    if image_path and os.path.exists(
        image_path
    ):

        image_status = (
            f"🟢 تصویر محلی: {image_path}"
        )

    elif image_file_id:

        image_status = (
            "🟡 تصویر روبیکا ثبت شده است"
        )

    elif image_path:

        image_status = (
            "⚠️ مسیر تصویر ثبت شده ولی فایل پیدا نشد:\n"
            f"{image_path}"
        )

    else:

        image_status = (
            "🔴 تصویر ثبت نشده است"
        )

    text = (
        "📋 مشخصات کامل محصول\n\n"

        "━━━━━━━━━━━━━━━━━━\n"

        f"🔖 کد محصول: {code}\n\n"

        f"🔹 نام محصول:\n{name}\n\n"

        f"📝 توضیحات:\n{description}\n\n"

        f"💰 قیمت فعلی:\n"
        f"{format_price(price)} تومان\n\n"

        f"⏱ زمان تقریبی ارسال:\n"
        f"{delivery_days} روز کاری\n\n"

        f"📌 وضعیت محصول:\n"
        f"{status}\n\n"

        f"🖼 وضعیت تصویر:\n"
        f"{image_status}\n\n"

        "━━━━━━━━━━━━━━━━━━\n"

        "📊 آمار محصول\n\n"

        f"🛒 تعداد کل سفارش‌ها: {total_orders}\n"

        f"🔥 سفارش‌های {RECENT_ORDER_DAYS} روز اخیر: "
        f"{recent_orders}\n"

        f"⭐ میانگین امتیاز: {average_rating} / 5\n"

        f"💬 تعداد امتیاز/نظر: {review_count}\n"

        "━━━━━━━━━━━━━━━━━━"
    )

    await bot.send_message(
        chat_id=chat_id,
        text=text,
        chat_keypad=admin_products_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# Excel
# =========================================================

async def export_orders_to_excel(chat_id):

    if not OPENPYXL_AVAILABLE:

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ کتابخانه openpyxl در محیط فعلی نصب نیست.\n"
                "ابتدا openpyxl را نصب کنید."
            )
        )

        return

    cursor.execute("""
        SELECT
            o.id,
            o.created_at,
            o.product,
            o.name,
            o.phone,
            o.address,
            o.status,
            o.delivered_at,
            o.accepted_at,
            o.customer_acceptance,
            r.rating,
            r.comment,
            r.created_at,
            rr.request_type,
            rr.reason,
            rr.status,
            rr.created_at,
            rr.decided_at,
            rr.admin_note
        FROM orders o

        LEFT JOIN reviews r
            ON o.id = r.order_id

        LEFT JOIN return_requests rr
            ON o.id = rr.order_id

        ORDER BY o.id DESC
    """)

    orders = cursor.fetchall()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Orders"

    headers = [
        "شماره سفارش",
        "تاریخ ثبت",
        "محصول",
        "نام مشتری",
        "شماره تماس",
        "آدرس",
        "وضعیت",
        "تاریخ تحویل",
        "تاریخ تأیید دریافت",
        "تأیید مشتری",
        "⭐ امتیاز مشتری",
        "💬 نظر مشتری",
        "تاریخ ثبت نظر",
        "نوع درخواست تعویض/مرجوعی",
        "دلیل درخواست",
        "وضعیت درخواست",
        "تاریخ درخواست",
        "تاریخ تصمیم مدیر",
        "یادداشت مدیر"
    ]

    sheet.append(headers)

    for cell in sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for order in orders:
        sheet.append(order)

    widths = {
        "A": 14,
        "B": 20,
        "C": 14,
        "D": 25,
        "E": 18,
        "F": 45,
        "G": 20,
        "H": 20,
        "I": 25,
        "J": 50,
        "K": 18,
        "L": 50,
        "M": 20,
        "N": 25,
        "O": 50,
        "P": 25,
        "Q": 20,
        "R": 20,
        "S": 50
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                horizontal="right",
                vertical="center",
                wrap_text=True
            )

    sheet.freeze_panes = "A2"

    # =====================================================
    # Summary
    # =====================================================

    summary = workbook.create_sheet(
        "Summary"
    )

    total_orders = len(orders)

    cursor.execute("""
        SELECT status, COUNT(*)
        FROM orders
        GROUP BY status
    """)

    status_rows = cursor.fetchall()

    summary.append([
        "گزارش کلی فروشگاه"
    ])

    summary["A1"].font = Font(
        bold=True
    )

    summary.append([
        "تعداد کل سفارش‌ها",
        total_orders
    ])

    summary.append([])

    summary.append([
        "وضعیت",
        "تعداد"
    ])

    for status, count in status_rows:

        summary.append([
            status,
            count
        ])

    cursor.execute("""
        SELECT
            AVG(rating),
            COUNT(*)
        FROM reviews
    """)

    rating_summary = cursor.fetchone()

    average_rating = (
        rating_summary[0]
        if rating_summary
        else None
    )

    rating_count = (
        rating_summary[1]
        if rating_summary
        else 0
    )

    if average_rating is None:
        average_rating = 0

    average_rating = round(
        average_rating,
        1
    )

    summary.append([])

    summary.append([
        "میانگین امتیاز مشتریان",
        average_rating
    ])

    summary.append([
        "تعداد امتیازهای ثبت‌شده",
        rating_count
    ])

    # آمار درخواست‌ها

    cursor.execute("""
        SELECT
            status,
            COUNT(*)
        FROM return_requests
        GROUP BY status
    """)

    return_rows = cursor.fetchall()

    summary.append([])

    summary.append([
        "درخواست‌های تعویض/مرجوعی"
    ])

    summary["A10"].font = Font(
        bold=True
    )

    summary.append([
        "وضعیت درخواست",
        "تعداد"
    ])

    for status, count in return_rows:

        summary.append([
            return_status_fa(status),
            count
        ])

    summary.column_dimensions[
        "A"
    ].width = 40

    summary.column_dimensions[
        "B"
    ].width = 25

    for row in summary.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                horizontal="right",
                vertical="center",
                wrap_text=True
            )

    filename = (
        "orders_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".xlsx"
    )

    file_path = os.path.abspath(
        filename
    )

    workbook.save(
        file_path
    )

    await bot.send_document(
        chat_id=chat_id,
        path=file_path,
        text="📊 فایل Excel سفارش‌ها آماده است."
    )

    try:
        os.remove(file_path)
    except Exception:
        pass


# =========================================================
# اعلان مدیر سفارش
# =========================================================

async def notify_admin(order_id):

    cursor.execute("""
        SELECT
            product,
            name,
            phone,
            address,
            status,
            created_at
        FROM orders
        WHERE id = ?
    """, (order_id,))

    order = cursor.fetchone()

    if not order:
        return

    (
        product,
        name,
        phone,
        address,
        status,
        created_at
    ) = order

    text = (
        "🔔 سفارش جدید فروشگاه\n\n"

        f"🧾 شماره سفارش: #{order_id}\n"

        f"📦 محصول: {product}\n\n"

        f"👤 نام مشتری: {name}\n"

        f"📱 شماره تماس: {phone}\n"

        f"📍 آدرس: {address}\n\n"

        f"📌 وضعیت: {status_fa(status)}\n"

        f"🕐 تاریخ ثبت: {created_at}\n\n"

        "لطفاً وضعیت سفارش را مشخص کنید."
    )

    await bot.send_message(
        chat_id=ADMIN_CHAT_ID,
        text=text,
        chat_keypad=admin_order_menu(
            order_id
        ),
        chat_keypad_type="New"
    )


# =========================================================
# اعلان وضعیت سفارش به مشتری
# =========================================================

async def notify_customer_status(
    order_id,
    status
):

    cursor.execute("""
        SELECT
            chat_id,
            name,
            product
        FROM orders
        WHERE id = ?
    """, (order_id,))

    result = cursor.fetchone()

    if not result:
        return

    (
        chat_id,
        name,
        product
    ) = result

    if not chat_id:
        return

    text = (
        f"سلام {name} عزیز 🌷\n\n"

        f"وضعیت سفارش #{order_id} تغییر کرد.\n\n"

        f"📦 محصول: {product}\n"

        f"📌 وضعیت جدید: "
        f"{status_fa(status)}"
    )

    await bot.send_message(
        chat_id=chat_id,
        text=text
    )


# =========================================================
# اعلان درخواست تعویض / مرجوعی به مدیر
# =========================================================

async def notify_admin_return_request(
    request_id
):

    cursor.execute("""
        SELECT
            rr.id,
            rr.order_id,
            rr.user_id,
            rr.product,
            rr.request_type,
            rr.reason,
            rr.status,
            rr.created_at,

            o.name,
            o.phone,
            o.address,
            o.accepted_at

        FROM return_requests rr

        LEFT JOIN orders o
            ON rr.order_id = o.id

        WHERE rr.id = ?
    """, (request_id,))

    result = cursor.fetchone()

    if not result:
        return

    (
        request_id,
        order_id,
        user_id,
        product,
        request_type,
        reason,
        request_status,
        created_at,
        customer_name,
        phone,
        address,
        accepted_at
    ) = result

    deadline = get_return_deadline(
        accepted_at
    )

    deadline_text = (
        deadline.strftime(
            "%Y-%m-%d %H:%M:%S"
        )
        if deadline
        else "نامشخص"
    )

    text = (
        "🔔 درخواست جدید تعویض / مرجوعی\n\n"

        f"🧾 شماره درخواست: #{request_id}\n"

        f"📦 شماره سفارش: #{order_id}\n"

        f"📦 محصول: {product}\n\n"

        f"👤 نام مشتری: {customer_name}\n"

        f"📱 شماره تماس: {phone}\n"

        f"📍 آدرس: {address}\n\n"

        f"🔄 نوع درخواست: {request_type}\n\n"

        f"📝 دلیل مشتری:\n"
        f"{reason}\n\n"

        f"🕐 زمان ثبت درخواست:\n"
        f"{created_at}\n\n"

        f"⏰ پایان مهلت درخواست:\n"
        f"{deadline_text}\n\n"

        "لطفاً درخواست را بررسی کنید."
    )

    await bot.send_message(
        chat_id=ADMIN_CHAT_ID,
        text=text,
        chat_keypad=admin_return_menu(
            request_id
        ),
        chat_keypad_type="New"
    )


# =========================================================
# اعلان نتیجه درخواست به مشتری
# =========================================================

async def notify_customer_return_decision(
    request_id
):

    cursor.execute("""
        SELECT
            rr.order_id,
            rr.user_id,
            rr.request_type,
            rr.status,
            rr.admin_note,
            o.chat_id,
            o.name,
            o.product
        FROM return_requests rr

        LEFT JOIN orders o
            ON rr.order_id = o.id

        WHERE rr.id = ?
    """, (request_id,))

    result = cursor.fetchone()

    if not result:
        return

    (
        order_id,
        user_id,
        request_type,
        request_status,
        admin_note,
        customer_chat_id,
        customer_name,
        product
    ) = result

    if not customer_chat_id:
        return

    if request_status == "تأیید شده":

        text = (
            f"سلام {customer_name} عزیز 🌷\n\n"

            f"درخواست {request_type} شما برای "
            f"سفارش #{order_id} بررسی شد.\n\n"

            "✅ درخواست شما توسط فروشگاه تأیید شد.\n\n"

            f"📦 محصول: {product}\n"
        )

    elif request_status == "رد شده":

        text = (
            f"سلام {customer_name} عزیز 🌷\n\n"

            f"درخواست {request_type} شما برای "
            f"سفارش #{order_id} بررسی شد.\n\n"

            "❌ درخواست شما توسط فروشگاه رد شد.\n\n"

            f"📦 محصول: {product}\n"
        )

    else:

        text = (
            f"سلام {customer_name} عزیز 🌷\n\n"

            f"وضعیت درخواست شما برای سفارش "
            f"#{order_id} تغییر کرد.\n\n"

            f"📌 وضعیت: "
            f"{return_status_fa(request_status)}"
        )

    if admin_note:

        text += (
            "\n\n"
            "📝 توضیحات فروشگاه:\n"
            f"{admin_note}"
        )

    await bot.send_message(
        chat_id=customer_chat_id,
        text=text
    )


# =========================================================
# پیام‌های متنی
# =========================================================

@bot.on_message()
async def handle_message(
    bot_instance: Robot,
    message: Message
):

    text = (
        getattr(
            message,
            "text",
            ""
        )
        or ""
    )

    user_id = get_user_id(
        message
    )

    chat_id = get_chat_id(
        message
    )

    # =====================================================
    # مدیر
    # =====================================================

    if chat_id == ADMIN_CHAT_ID:

        current_admin_state = (
            admin_states.get(chat_id)
        )

        if (
            current_admin_state
            and text.strip()
            in [
                "لغو",
                "لغو عملیات",
                "cancel"
            ]
        ):

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text="❌ عملیات لغو شد.",
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # تغییر قیمت
        # -------------------------------------------------

        if current_admin_state == "waiting_price":

            new_price = normalize_digits(
                text.strip()
            )

            new_price = (
                new_price
                .replace(",", "")
                .replace("٬", "")
                .replace(" ", "")
            )

            if not new_price.isdigit():

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ قیمت واردشده معتبر نیست.\n\n"
                        "لطفاً فقط عدد وارد کنید.\n\n"
                        "مثال:\n"
                        "1500000"
                    )
                )

                return

            try:

                price_number = int(
                    new_price
                )

                if price_number <= 0:
                    raise ValueError

            except ValueError:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ قیمت باید یک عدد بزرگ‌تر از صفر باشد.\n\n"
                        "مثال:\n"
                        "1500000"
                    )
                )

                return

            cursor.execute("""
                UPDATE products
                SET price = ?
                WHERE code = ?
            """, (
                str(price_number),
                PRODUCT_CODE
            ))

            conn.commit()

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ قیمت محصول با موفقیت تغییر کرد.\n\n"

                    f"📦 محصول: {PRODUCT_CODE}\n"

                    f"💰 قیمت جدید: "
                    f"{format_price(price_number)} تومان"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # تغییر نام
        # -------------------------------------------------

        if current_admin_state == "waiting_name":

            new_name = text.strip()

            if not new_name:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ نام محصول نمی‌تواند خالی باشد.\n\n"
                        "لطفاً نام جدید محصول را وارد کنید."
                    )
                )

                return

            if len(new_name) > 150:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ نام محصول بیش از حد طولانی است.\n\n"
                        "حداکثر 150 کاراکتر."
                    )
                )

                return

            cursor.execute("""
                UPDATE products
                SET name = ?
                WHERE code = ?
            """, (
                new_name,
                PRODUCT_CODE
            ))

            conn.commit()

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ نام محصول با موفقیت تغییر کرد.\n\n"

                    f"📦 کد محصول: {PRODUCT_CODE}\n"

                    f"🔹 نام جدید:\n{new_name}"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # تغییر توضیحات
        # -------------------------------------------------

        if current_admin_state == "waiting_description":

            new_description = text.strip()

            if not new_description:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ توضیحات محصول نمی‌تواند خالی باشد.\n\n"
                        "لطفاً توضیحات جدید محصول را وارد کنید."
                    )
                )

                return

            if len(new_description) > 2000:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ توضیحات محصول بیش از حد طولانی است.\n\n"
                        "حداکثر 2000 کاراکتر."
                    )
                )

                return

            cursor.execute("""
                UPDATE products
                SET description = ?
                WHERE code = ?
            """, (
                new_description,
                PRODUCT_CODE
            ))

            conn.commit()

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ توضیحات محصول با موفقیت تغییر کرد.\n\n"

                    f"📦 کد محصول: {PRODUCT_CODE}\n\n"

                    f"📝 توضیحات جدید:\n"
                    f"{new_description}"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # تغییر زمان ارسال
        # -------------------------------------------------

        if current_admin_state == "waiting_delivery_days":

            new_delivery_days = text.strip()

            if not new_delivery_days:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ زمان ارسال نمی‌تواند خالی باشد.\n\n"

                        "مثال:\n"
                        "1 تا 3\n\n"

                        "یا:\n"
                        "2 تا 4"
                    )
                )

                return

            if len(new_delivery_days) > 100:

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ متن زمان ارسال بیش از حد طولانی است."
                    )
                )

                return

            cursor.execute("""
                UPDATE products
                SET delivery_days = ?
                WHERE code = ?
            """, (
                new_delivery_days,
                PRODUCT_CODE
            ))

            conn.commit()

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ زمان ارسال با موفقیت تغییر کرد.\n\n"

                    f"📦 محصول: {PRODUCT_CODE}\n"

                    f"⏱ زمان جدید ارسال:\n"
                    f"{new_delivery_days} روز کاری"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # دستورات مدیر
        # -------------------------------------------------

        if text == "/export":

            await export_orders_to_excel(
                chat_id
            )

            return

        if text == "/admin":

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text="👨‍💼 پنل مدیریت فروشگاه",
                chat_keypad=admin_menu(),
                chat_keypad_type="New"
            )

            return

    # =====================================================
    # /start
    # =====================================================

    if text == "/start":

        customers.pop(
            user_id,
            None
        )

        rating_comments.pop(
            user_id,
            None
        )

        return_request_states.pop(
            user_id,
            None
        )

        if chat_id == ADMIN_CHAT_ID:

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "👨‍💼 سلام مدیر محترم\n\n"
                    "به پنل مدیریت فروشگاه خوش آمدید."
                ),
                chat_keypad=admin_menu(),
                chat_keypad_type="New"
            )

        else:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "🏪 به فروشگاه آنلاین شاپ خوش آمدید.\n\n"

                    "خریدی آسان و مطمئن برای کلیه همشهریان "
                    "و هموطنان\n\n"

                    "لطفاً یکی از گزینه‌های زیر را انتخاب کنید."
                ),
                chat_keypad=main_menu(),
                chat_keypad_type="New"
            )

        return

    # =====================================================
    # درخواست تعویض / مرجوعی در انتظار دریافت دلیل
    # =====================================================

    if user_id in return_request_states:

        state = return_request_states[
            user_id
        ]

        order_id = state.get(
            "order_id"
        )

        request_type = state.get(
            "request_type"
        )

        if text.strip() in [
            "لغو",
            "لغو عملیات",
            "cancel"
        ]:

            return_request_states.pop(
                user_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text="❌ درخواست تعویض/مرجوعی لغو شد.",
                chat_keypad=main_menu(),
                chat_keypad_type="New"
            )

            return

        reason = text.strip()

        if not reason:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ دلیل درخواست نمی‌تواند خالی باشد.\n\n"
                    "لطفاً دلیل تعویض یا مرجوعی را بنویسید."
                )
            )

            return

        # بررسی مجدد مهلت
        cursor.execute("""
            SELECT
                user_id,
                product,
                accepted_at
            FROM orders
            WHERE id = ?
        """, (order_id,))

        order_result = cursor.fetchone()

        if not order_result:

            return_request_states.pop(
                user_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text="❌ سفارش پیدا نشد."
            )

            return

        (
            order_user_id,
            product_code,
            accepted_at
        ) = order_result

        if order_user_id != user_id:

            return_request_states.pop(
                user_id,
                None
            )

            return

        if not is_return_window_open(
            accepted_at
        ):

            deadline = get_return_deadline(
                accepted_at
            )

            deadline_text = (
                deadline.strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                if deadline
                else "نامشخص"
            )

            return_request_states.pop(
                user_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⏰ مهلت درخواست تعویض/مرجوعی "
                    "این سفارش به پایان رسیده است.\n\n"

                    f"🧾 سفارش: #{order_id}\n"

                    f"⏰ پایان مهلت: {deadline_text}"
                ),
                chat_keypad=main_menu(),
                chat_keypad_type="New"
            )

            return

        # بررسی درخواست قبلی
        cursor.execute("""
            SELECT
                id,
                status
            FROM return_requests
            WHERE order_id = ?
        """, (order_id,))

        existing_request = cursor.fetchone()

        if existing_request:

            return_request_states.pop(
                user_id,
                None
            )

            existing_request_id, existing_status = (
                existing_request
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ برای این سفارش قبلاً درخواست "
                    "تعویض/مرجوعی ثبت شده است.\n\n"

                    f"🧾 شماره درخواست: #{existing_request_id}\n"

                    f"📌 وضعیت: "
                    f"{return_status_fa(existing_status)}"
                ),
                chat_keypad=main_menu(),
                chat_keypad_type="New"
            )

            return

        created_at = now_string()

        cursor.execute("""
            INSERT INTO return_requests
            (
                order_id,
                user_id,
                product,
                request_type,
                reason,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            order_id,
            user_id,
            product_code,
            request_type,
            reason,
            "در انتظار بررسی",
            created_at
        ))

        conn.commit()

        request_id = cursor.lastrowid

        return_request_states.pop(
            user_id,
            None
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "✅ درخواست شما با موفقیت ثبت شد.\n\n"

                f"🧾 شماره درخواست: #{request_id}\n"

                f"📦 سفارش: #{order_id}\n"

                f"🔄 نوع درخواست: {request_type}\n\n"

                "درخواست برای مدیر فروشگاه ارسال شد.\n"

                "پس از بررسی، نتیجه برای شما ارسال خواهد شد."
            ),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        await notify_admin_return_request(
            request_id
        )

        return

    # =====================================================
    # ثبت سفارش مشتری
    # =====================================================

    if user_id in customers:

        state = customers[
            user_id
        ]

        step = state.get(
            "step"
        )

        if step == "name":

            state["name"] = text.strip()

            state["step"] = "phone"

            await bot.send_message(
                chat_id=chat_id,
                text="📱 لطفاً شماره تماس خود را وارد کنید:"
            )

            return

        if step == "phone":

            state["phone"] = text.strip()

            state["step"] = "address"

            await bot.send_message(
                chat_id=chat_id,
                text="📍 لطفاً آدرس کامل خود را وارد کنید:"
            )

            return

        if step == "address":

            state["address"] = text.strip()

            state["step"] = "confirm"

            product = get_product()

            if not product:

                customers.pop(
                    user_id,
                    None
                )

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "❌ محصول پیدا نشد. "
                        "لطفاً دوباره تلاش کنید."
                    ),
                    chat_keypad=main_menu(),
                    chat_keypad_type="New"
                )

                return

            (
                code,
                name,
                description,
                price,
                delivery_days,
                image_path,
                image_file_id
            ) = product

            summary = (
                "🧾 خلاصه سفارش\n\n"

                f"📦 محصول: {code} - {name}\n"

                f"💰 قیمت: {format_price(price)} تومان\n\n"

                f"👤 نام: {state['name']}\n"

                f"📱 تلفن: {state['phone']}\n"

                f"📍 آدرس: {state['address']}\n\n"

                f"🚚 ارسال رایگان فقط در محدوده "
                f"{FREE_DELIVERY_CITY}\n"

                f"⏱ زمان تقریبی ارسال: "
                f"{delivery_days} روز کاری\n"

                "💵 پرداخت درب منزل پس از رویت و تحویل کالا\n\n"

                "آیا اطلاعات سفارش صحیح است؟"
            )

            await bot.send_message(
                chat_id=chat_id,
                text=summary,
                chat_keypad=confirm_menu(),
                chat_keypad_type="New"
            )

            return


# =========================================================
# دریافت تصویر جدید مدیر
# =========================================================

@bot.on_message_file()
async def handle_admin_image(
    bot_instance: Robot,
    message: Message
):

    chat_id = get_chat_id(
        message
    )

    if chat_id != ADMIN_CHAT_ID:
        return

    if admin_states.get(
        chat_id
    ) != "waiting_image":

        return

    file_obj = getattr(
        message,
        "file",
        None
    )

    if not file_obj:

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ اطلاعات فایل تصویر دریافت نشد.\n\n"
                "لطفاً دوباره تصویر را به صورت Photo ارسال کنید."
            )
        )

        return

    file_id = None

    for attr in (
        "file_id",
        "id",
        "fileId"
    ):

        try:

            value = getattr(
                file_obj,
                attr,
                None
            )

            if value:

                file_id = str(
                    value
                )

                break

        except Exception:
            pass

    if (
        not file_id
        and isinstance(
            file_obj,
            dict
        )
    ):

        for key in (
            "file_id",
            "id",
            "fileId"
        ):

            value = file_obj.get(
                key
            )

            if value:

                file_id = str(
                    value
                )

                break

    if not file_id:

        raw_data = getattr(
            message,
            "raw_data",
            None
        ) or {}

        def find_file_id(obj):

            if isinstance(
                obj,
                dict
            ):

                for key in (
                    "file_id",
                    "fileId"
                ):

                    value = obj.get(
                        key
                    )

                    if value:
                        return str(
                            value
                        )

                for value in obj.values():

                    found = find_file_id(
                        value
                    )

                    if found:
                        return found

            elif isinstance(
                obj,
                list
            ):

                for item in obj:

                    found = find_file_id(
                        item
                    )

                    if found:
                        return found

            return None

        file_id = find_file_id(
            raw_data
        )

    if not file_id:

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ شناسه تصویر از روبیکا دریافت نشد.\n\n"
                "لطفاً دوباره تصویر را ارسال کنید."
            )
        )

        return

    temp_path = (
        PRODUCT_IMAGE_PATH
        + ".new"
    )

    final_path = PRODUCT_IMAGE_PATH

    try:

        if os.path.exists(
            temp_path
        ):

            os.remove(
                temp_path
            )

    except Exception:
        pass

    try:

        print(
            "DOWNLOADING RUBIKA IMAGE..."
        )

        result = await bot.download(
            file_id=file_id,
            save_as=temp_path,
            verbose=False
        )

        print(
            "RUBIKA IMAGE DOWNLOAD RESULT:",
            result
        )

    except Exception as error:

        print(
            "RUBIKA IMAGE DOWNLOAD ERROR:",
            repr(error)
        )

        try:

            if os.path.exists(
                temp_path
            ):

                os.remove(
                    temp_path
                )

        except Exception:
            pass

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ دانلود تصویر از روبیکا انجام نشد.\n\n"
                "لطفاً دوباره تصویر را ارسال کنید.\n\n"
                "خطای فنی در کنسول ثبت شده است."
            )
        )

        return

    if not os.path.exists(
        temp_path
    ):

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ تصویر دریافت شد اما فایل روی سیستم ذخیره نشد.\n\n"
                "لطفاً دوباره تلاش کنید."
            )
        )

        return

    try:

        file_size = os.path.getsize(
            temp_path
        )

    except Exception:

        file_size = 0

    if file_size <= 0:

        try:

            os.remove(
                temp_path
            )

        except Exception:
            pass

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ فایل تصویر خالی است و قابل استفاده نیست.\n\n"
                "لطفاً دوباره تصویر را ارسال کنید."
            )
        )

        return

    try:

        if os.path.exists(
            final_path
        ):

            os.remove(
                final_path
            )

        os.replace(
            temp_path,
            final_path
        )

    except Exception as error:

        print(
            "IMAGE REPLACE ERROR:",
            repr(error)
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ ذخیره نهایی تصویر انجام نشد.\n\n"
                "لطفاً دوباره تلاش کنید."
            )
        )

        return

    cursor.execute("""
        UPDATE products
        SET
            image_path = ?,
            image_file_id = NULL
        WHERE code = ?
    """, (
        final_path,
        PRODUCT_CODE
    ))

    conn.commit()

    admin_states.pop(
        chat_id,
        None
    )

    await bot.send_message(
        chat_id=chat_id,
        text=(
            "✅ تصویر محصول با موفقیت تغییر کرد.\n\n"

            f"📦 محصول: {PRODUCT_CODE}\n"

            "🖼 تصویر جدید روی سیستم ذخیره شد.\n"

            f"📁 فایل: {final_path}\n"

            f"📦 حجم فایل: {file_size:,} بایت\n\n"

            "از این پس مشتریان تصویر جدید را مشاهده خواهند کرد."
        ),
        chat_keypad=admin_products_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# Callback
# =========================================================

@bot.on_callback()
async def handle_callback(
    bot_instance: Robot,
    message: Message
):

    button_id = get_button_id(
        message
    )

    if not button_id:
        return

    user_id = get_user_id(
        message
    )

    chat_id = get_chat_id(
        message
    )

    # =====================================================
    # مدیریت
    # =====================================================

    if button_id.startswith(
        "admin_"
    ):

        if not is_admin_message(
            message
        ):
            return

        # -------------------------------------------------
        # بازگشت
        # -------------------------------------------------

        if button_id == "admin_back":

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text="👨‍💼 پنل مدیریت فروشگاه",
                chat_keypad=admin_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # مدیریت محصولات
        # -------------------------------------------------

        if button_id == "admin_products":

            admin_states.pop(
                chat_id,
                None
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "📦 مدیریت محصولات\n\n"
                    "محصول موردنظر را انتخاب کنید:"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # مشخصات کامل
        # -------------------------------------------------

        if button_id == "admin_product_info_P-001":

            admin_states.pop(
                chat_id,
                None
            )

            await show_admin_product_info(
                chat_id
            )

            return

        # -------------------------------------------------
        # تغییر قیمت
        # -------------------------------------------------

        if button_id == "admin_price_P-001":

            cursor.execute("""
                SELECT price
                FROM products
                WHERE code = ?
            """, (PRODUCT_CODE,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            current_price = result[0]

            admin_states[
                chat_id
            ] = "waiting_price"

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "💰 تغییر قیمت محصول P-001\n\n"

                    f"قیمت فعلی:\n"
                    f"{format_price(current_price)} تومان\n\n"

                    "💵 لطفاً قیمت جدید را فقط به صورت عدد وارد کنید.\n\n"

                    "مثال:\n"
                    "1500000\n\n"

                    "برای لغو عملیات بنویسید:\n"
                    "لغو"
                )
            )

            return

        # -------------------------------------------------
        # تغییر نام
        # -------------------------------------------------

        if button_id == "admin_name_P-001":

            cursor.execute("""
                SELECT name
                FROM products
                WHERE code = ?
            """, (PRODUCT_CODE,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            current_name = result[0]

            admin_states[
                chat_id
            ] = "waiting_name"

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "📝 تغییر نام محصول P-001\n\n"

                    f"نام فعلی:\n"
                    f"{current_name}\n\n"

                    "لطفاً نام جدید محصول را وارد کنید.\n\n"

                    "برای لغو عملیات بنویسید:\n"
                    "لغو"
                )
            )

            return

        # -------------------------------------------------
        # تغییر توضیحات
        # -------------------------------------------------

        if button_id == "admin_description_P-001":

            cursor.execute("""
                SELECT description
                FROM products
                WHERE code = ?
            """, (PRODUCT_CODE,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            current_description = result[0]

            admin_states[
                chat_id
            ] = "waiting_description"

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "📄 تغییر توضیحات محصول P-001\n\n"

                    f"توضیحات فعلی:\n"
                    f"{current_description}\n\n"

                    "لطفاً توضیحات جدید محصول را در یک پیام ارسال کنید.\n\n"

                    "برای لغو عملیات بنویسید:\n"
                    "لغو"
                )
            )

            return

        # -------------------------------------------------
        # تغییر تصویر
        # -------------------------------------------------

        if button_id == "admin_image_P-001":

            product = get_product_admin()

            if not product:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            (
                code,
                name,
                description,
                price,
                delivery_days,
                image_path,
                image_file_id,
                active
            ) = product

            admin_states[
                chat_id
            ] = "waiting_image"

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "🖼 تغییر تصویر محصول P-001\n\n"

                    f"📦 محصول: {name}\n\n"

                    "لطفاً تصویر جدید محصول را همین الان "
                    "به صورت مستقیم به ربات ارسال کنید.\n\n"

                    "⚠️ تصویر را به صورت Photo ارسال کنید، "
                    "نه به صورت فایل Document.\n\n"

                    "تصویر دریافت می‌شود و روی سیستم فروشگاه "
                    "ذخیره خواهد شد.\n\n"

                    "پس از دریافت تصویر، به صورت خودکار ثبت می‌شود.\n\n"

                    "برای لغو عملیات بنویسید:\n"
                    "لغو"
                )
            )

            return

        # -------------------------------------------------
        # تغییر زمان
        # -------------------------------------------------

        if button_id == "admin_delivery_P-001":

            cursor.execute("""
                SELECT delivery_days
                FROM products
                WHERE code = ?
            """, (PRODUCT_CODE,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            current_delivery = result[0]

            admin_states[
                chat_id
            ] = "waiting_delivery_days"

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⏱ تغییر زمان ارسال P-001\n\n"

                    f"زمان فعلی:\n"
                    f"{current_delivery} روز کاری\n\n"

                    "زمان جدید را وارد کنید.\n\n"

                    "مثال:\n"
                    "1 تا 3\n\n"

                    "یا:\n"
                    "2 تا 4\n\n"

                    "برای لغو عملیات بنویسید:\n"
                    "لغو"
                )
            )

            return

        # -------------------------------------------------
        # فعال / غیرفعال
        # -------------------------------------------------

        if button_id == "admin_toggle_P-001":

            product = get_product_admin()

            if not product:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ محصول P-001 پیدا نشد."
                )

                return

            active = product[7]

            if active:

                new_active = 0

                status_text = "🔴 غیرفعال"

            else:

                new_active = 1

                status_text = "🟢 فعال"

            cursor.execute("""
                UPDATE products
                SET active = ?
                WHERE code = ?
            """, (
                new_active,
                PRODUCT_CODE
            ))

            conn.commit()

            admin_states.pop(
                chat_id,
                None
            )

            if new_active:

                customer_message = (
                    "محصول دوباره در فروشگاه "
                    "برای مشتریان نمایش داده می‌شود."
                )

            else:

                customer_message = (
                    "محصول دیگر در بخش محصولات "
                    "مشتریان نمایش داده نمی‌شود."
                )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ وضعیت محصول تغییر کرد.\n\n"

                    f"📦 محصول: {PRODUCT_CODE}\n"

                    f"📌 وضعیت جدید: {status_text}\n\n"

                    f"{customer_message}"
                ),
                chat_keypad=admin_products_menu(),
                chat_keypad_type="New"
            )

            return

        # -------------------------------------------------
        # Excel
        # -------------------------------------------------

        if button_id == "admin_export_orders":

            admin_states.pop(
                chat_id,
                None
            )

            await export_orders_to_excel(
                chat_id
            )

            return

        # -------------------------------------------------
        # سفارش‌ها
        # -------------------------------------------------

        if button_id == "admin_orders":

            admin_states.pop(
                chat_id,
                None
            )

            await show_admin_orders(
                chat_id
            )

            return

        # -------------------------------------------------
        # تأیید سفارش
        # -------------------------------------------------

        if button_id.startswith(
            "admin_confirm_"
        ):

            try:

                order_id = int(
                    button_id.replace(
                        "admin_confirm_",
                        ""
                    )
                )

            except ValueError:

                return

            cursor.execute("""
                SELECT status
                FROM orders
                WHERE id = ?
            """, (order_id,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ سفارش پیدا نشد."
                )

                return

            cursor.execute("""
                UPDATE orders
                SET status = ?
                WHERE id = ?
            """, (
                "تأیید شده",
                order_id
            ))

            conn.commit()

            await bot.send_message(
                chat_id=chat_id,
                text=f"✅ سفارش #{order_id} تأیید شد."
            )

            await notify_customer_status(
                order_id,
                "تأیید شده"
            )

            return

        # -------------------------------------------------
        # لغو سفارش
        # -------------------------------------------------

        if button_id.startswith(
            "admin_cancel_"
        ):

            try:

                order_id = int(
                    button_id.replace(
                        "admin_cancel_",
                        ""
                    )
                )

            except ValueError:

                return

            cursor.execute("""
                UPDATE orders
                SET status = ?
                WHERE id = ?
            """, (
                "لغو شده",
                order_id
            ))

            conn.commit()

            await bot.send_message(
                chat_id=chat_id,
                text=f"❌ سفارش #{order_id} لغو شد."
            )

            await notify_customer_status(
                order_id,
                "لغو شده"
            )

            return

        # -------------------------------------------------
        # تحویل سفارش
        # -------------------------------------------------

        if button_id.startswith(
            "admin_deliver_"
        ):

            try:

                order_id = int(
                    button_id.replace(
                        "admin_deliver_",
                        ""
                    )
                )

            except ValueError:

                return

            cursor.execute("""
                SELECT
                    chat_id,
                    name,
                    product
                FROM orders
                WHERE id = ?
            """, (order_id,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ سفارش پیدا نشد."
                )

                return

            (
                customer_chat_id,
                customer_name,
                product
            ) = result

            cursor.execute("""
                UPDATE orders
                SET
                    status = ?,
                    delivered_at = ?
                WHERE id = ?
            """, (
                "تحویل شده",
                now_string(),
                order_id
            ))

            conn.commit()

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    f"🚚 سفارش #{order_id} "
                    "به عنوان تحویل شده ثبت شد."
                )
            )

            if customer_chat_id:

                await bot.send_message(
                    chat_id=customer_chat_id,
                    text=(
                        f"سلام {customer_name} عزیز 🌷\n\n"

                        f"📦 سفارش #{order_id} "
                        f"({product}) تحویل شما شده است.\n\n"

                        "لطفاً پس از رویت و بررسی کالا، "
                        "در صورت سالم بودن و مطابقت با سفارش، "
                        "دریافت آن را تأیید کنید."
                    ),
                    chat_keypad=customer_accept_menu(
                        order_id
                    ),
                    chat_keypad_type="New"
                )

            return

        # -------------------------------------------------
        # تأیید درخواست تعویض / مرجوعی
        # -------------------------------------------------

        if button_id.startswith(
            "admin_return_approve_"
        ):

            try:

                request_id = int(
                    button_id.replace(
                        "admin_return_approve_",
                        ""
                    )
                )

            except ValueError:

                return

            cursor.execute("""
                SELECT
                    id,
                    order_id,
                    status
                FROM return_requests
                WHERE id = ?
            """, (request_id,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ درخواست پیدا نشد."
                )

                return

            (
                request_id_db,
                order_id,
                current_status
            ) = result

            if current_status != "در انتظار بررسی":

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ این درخواست قبلاً بررسی شده است.\n\n"

                        f"📌 وضعیت فعلی: "
                        f"{return_status_fa(current_status)}"
                    )
                )

                return

            decision_time = now_string()

            cursor.execute("""
                UPDATE return_requests
                SET
                    status = ?,
                    decided_at = ?,
                    admin_note = ?
                WHERE id = ?
            """, (
                "تأیید شده",
                decision_time,
                "درخواست توسط مدیر تأیید شد.",
                request_id
            ))

            conn.commit()

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    f"✅ درخواست #{request_id} "
                    "تأیید شد.\n\n"

                    f"🧾 سفارش: #{order_id}"
                )
            )

            await notify_customer_return_decision(
                request_id
            )

            return

        # -------------------------------------------------
        # رد درخواست تعویض / مرجوعی
        # -------------------------------------------------

        if button_id.startswith(
            "admin_return_reject_"
        ):

            try:

                request_id = int(
                    button_id.replace(
                        "admin_return_reject_",
                        ""
                    )
                )

            except ValueError:

                return

            cursor.execute("""
                SELECT
                    id,
                    order_id,
                    status
                FROM return_requests
                WHERE id = ?
            """, (request_id,))

            result = cursor.fetchone()

            if not result:

                await bot.send_message(
                    chat_id=chat_id,
                    text="❌ درخواست پیدا نشد."
                )

                return

            (
                request_id_db,
                order_id,
                current_status
            ) = result

            if current_status != "در انتظار بررسی":

                await bot.send_message(
                    chat_id=chat_id,
                    text=(
                        "⚠️ این درخواست قبلاً بررسی شده است.\n\n"

                        f"📌 وضعیت فعلی: "
                        f"{return_status_fa(current_status)}"
                    )
                )

                return

            decision_time = now_string()

            cursor.execute("""
                UPDATE return_requests
                SET
                    status = ?,
                    decided_at = ?,
                    admin_note = ?
                WHERE id = ?
            """, (
                "رد شده",
                decision_time,
                "درخواست توسط مدیر رد شد.",
                request_id
            ))

            conn.commit()

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    f"❌ درخواست #{request_id} "
                    "رد شد.\n\n"

                    f"🧾 سفارش: #{order_id}"
                )
            )

            await notify_customer_return_decision(
                request_id
            )

            return

        return

    # =====================================================
    # منوی اصلی
    # =====================================================

    if button_id == "main_menu":

        customers.pop(
            user_id,
            None
        )

        return_request_states.pop(
            user_id,
            None
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "🏪 فروشگاه آنلاین شاپ\n\n"
                "لطفاً گزینه موردنظر را انتخاب کنید."
            ),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        return

    # =====================================================
    # نمایش محصول
    # =====================================================

    if button_id == "show_products":

        await show_product(
            chat_id
        )

        return

    # =====================================================
    # قوانین
    # =====================================================

    if button_id == "shop_policy":

        await bot.send_message(
            chat_id=chat_id,
            text=policy_text(),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        return

    # =====================================================
    # سفارش‌های مشتری
    # =====================================================

    if button_id == "my_orders":

        await show_my_orders(
            message
        )

        return

    # =====================================================
    # شروع سفارش
    # =====================================================

    if button_id == "order_product_P-001":

        product = get_product()

        if not product:

            await bot.send_message(
                chat_id=chat_id,
                text="❌ محصول موردنظر موجود نیست."
            )

            return

        customers[
            user_id
        ] = {
            "step": "name",
            "product": PRODUCT_CODE,
            "name": "",
            "phone": "",
            "address": ""
        }

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "🛒 ثبت سفارش\n\n"

                "برای ثبت سفارش ابتدا نام و نام خانوادگی "
                "خود را وارد کنید:"
            )
        )

        return

    # =====================================================
    # تأیید سفارش مشتری
    # =====================================================

    if button_id == "customer_confirm_order":

        if user_id not in customers:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ اطلاعات سفارش شما منقضی شده است. "
                    "لطفاً دوباره سفارش دهید."
                )
            )

            return

        state = customers[
            user_id
        ]

        if state.get(
            "step"
        ) != "confirm":

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ این سفارش در وضعیت قابل تأیید نیست."
                )
            )

            return

        product_code = state[
            "product"
        ]

        cursor.execute("""
            INSERT INTO orders
            (
                user_id,
                product,
                name,
                phone,
                address,
                status,
                chat_id,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            user_id,
            product_code,
            state["name"],
            state["phone"],
            state["address"],
            "جدید",
            chat_id,
            now_string()
        ))

        conn.commit()

        order_id = cursor.lastrowid

        customers.pop(
            user_id,
            None
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "✅ سفارش شما با موفقیت ثبت شد.\n\n"

                f"🧾 شماره سفارش: #{order_id}\n"

                "📌 وضعیت: 🆕 جدید\n\n"

                "اطلاعات سفارش برای مدیر فروشگاه ارسال شد."
            ),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        await notify_admin(
            order_id
        )

        return

    # =====================================================
    # لغو سفارش
    # =====================================================

    if button_id == "customer_cancel_order":

        customers.pop(
            user_id,
            None
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "❌ سفارش لغو شد.\n\n"

                "می‌توانید دوباره از منوی اصلی محصول موردنظر "
                "را انتخاب کنید."
            ),
            chat_keypad=main_menu(),
            chat_keypad_type="New"
        )

        return

    # =====================================================
    # تأیید دریافت کالا
    # =====================================================

    if button_id.startswith(
        "customer_accept_delivery_"
    ):

        try:

            order_id = int(
                button_id.replace(
                    "customer_accept_delivery_",
                    ""
                )
            )

        except ValueError:

            return

        cursor.execute("""
            SELECT
                user_id,
                status
            FROM orders
            WHERE id = ?
        """, (order_id,))

        result = cursor.fetchone()

        if not result:

            await bot.send_message(
                chat_id=chat_id,
                text="❌ سفارش پیدا نشد."
            )

            return

        (
            order_user_id,
            status
        ) = result

        if order_user_id != user_id:
            return

        if status != "تحویل شده":

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ این سفارش هنوز در وضعیت تحویل شده "
                    "قرار نگرفته است."
                )
            )

            return

        accepted_at = now_string()

        cursor.execute("""
            UPDATE orders
            SET
                status = ?,
                accepted_at = ?,
                customer_acceptance = ?
            WHERE id = ?
        """, (
            "تحویل و تأیید شده",
            accepted_at,
            "مشتری اعلام کرد کالا را رویت کرده و سالم و مطابق سفارش تحویل گرفته است.",
            order_id
        ))

        conn.commit()

        deadline = get_return_deadline(
            accepted_at
        )

        deadline_text = (
            deadline.strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            if deadline
            else "نامشخص"
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "✅ دریافت کالا ثبت شد.\n\n"

                "از خرید شما سپاسگزاریم 🌷\n\n"

                f"🔄 مهلت درخواست تعویض/مرجوعی "
                f"{RETURN_WINDOW_HOURS} ساعت است.\n\n"

                f"⏰ پایان مهلت:\n"
                f"{deadline_text}\n\n"

                "لطفاً تجربه خود از خرید را با ثبت امتیاز "
                "با ما به اشتراک بگذارید:"
            ),
            chat_keypad=rating_menu(
                order_id
            ),
            chat_keypad_type="New"
        )

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "🔄 در صورت وجود ایراد یا مغایرت، "
                "می‌توانید در مهلت تعیین‌شده درخواست "
                "تعویض یا مرجوعی ثبت کنید."
            ),
            chat_keypad=customer_return_menu(
                order_id
            ),
            chat_keypad_type="New"
        )

        await bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=(
                f"📦 مشتری سفارش #{order_id} را تحویل و تأیید کرد.\n\n"

                "مشتری اعلام کرده کالا را رویت کرده و سالم و "
                "مطابق سفارش تحویل گرفته است.\n\n"

                f"⏰ مهلت تعویض/مرجوعی تا:\n"
                f"{deadline_text}"
            )
        )

        return

    # =====================================================
    # شروع درخواست تعویض / مرجوعی
    # =====================================================

    if button_id.startswith(
        "customer_return_"
    ):

        try:

            order_id = int(
                button_id.replace(
                    "customer_return_",
                    ""
                )
            )

        except ValueError:

            return

        cursor.execute("""
            SELECT
                user_id,
                product,
                status,
                accepted_at
            FROM orders
            WHERE id = ?
        """, (order_id,))

        result = cursor.fetchone()

        if not result:

            await bot.send_message(
                chat_id=chat_id,
                text="❌ سفارش پیدا نشد."
            )

            return

        (
            order_user_id,
            product_code,
            status,
            accepted_at
        ) = result

        if order_user_id != user_id:
            return

        if status != "تحویل و تأیید شده":

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ ابتدا باید دریافت کالا را "
                    "تأیید کرده باشید."
                )
            )

            return

        deadline = get_return_deadline(
            accepted_at
        )

        if not deadline:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ زمان تأیید دریافت سفارش "
                    "قابل شناسایی نیست."
                )
            )

            return

        if datetime.now() > deadline:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⏰ مهلت درخواست تعویض/مرجوعی "
                    "این سفارش به پایان رسیده است.\n\n"

                    f"🧾 سفارش: #{order_id}\n"

                    f"⏰ پایان مهلت:\n"
                    f"{deadline.strftime('%Y-%m-%d %H:%M:%S')}"
                )
            )

            return

        cursor.execute("""
            SELECT
                id,
                status
            FROM return_requests
            WHERE order_id = ?
        """, (order_id,))

        existing_request = cursor.fetchone()

        if existing_request:

            request_id, request_status = (
                existing_request
            )

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⚠️ برای این سفارش قبلاً درخواست "
                    "تعویض/مرجوعی ثبت شده است.\n\n"

                    f"🧾 شماره درخواست: #{request_id}\n"

                    f"📌 وضعیت: "
                    f"{return_status_fa(request_status)}"
                )
            )

            return

        return_request_states[
            user_id
        ] = {
            "order_id": order_id,
            "request_type": "تعویض / مرجوعی"
        }

        await bot.send_message(
            chat_id=chat_id,
            text=(
                "🔄 درخواست تعویض / مرجوعی\n\n"

                f"🧾 شماره سفارش: #{order_id}\n\n"

                "لطفاً دلیل درخواست خود را در پیام بعدی "
                "به طور کامل بنویسید.\n\n"

                "مثلاً:\n"
                "کالا با سفارش من مغایرت دارد.\n\n"

                "یا:\n"
                "کالا هنگام تحویل دارای آسیب‌دیدگی بود.\n\n"

                "برای لغو عملیات بنویسید:\n"
                "لغو"
            )
        )

        return

    # =====================================================
    # امتیاز
    # =====================================================

    if button_id.startswith(
        "rating_"
    ):

        parts = button_id.split(
            "_"
        )

        if len(parts) != 3:
            return

        try:

            order_id = int(
                parts[1]
            )

            rating = int(
                parts[2]
            )

        except ValueError:

            return

        if rating < 1 or rating > 5:
            return

        cursor.execute("""
            SELECT
                user_id,
                product,
                status
            FROM orders
            WHERE id = ?
        """, (order_id,))

        result = cursor.fetchone()

        if not result:
            return

        (
            order_user_id,
            product_code,
            status
        ) = result

        if order_user_id != user_id:
            return

        if status != "تحویل و تأیید شده":
            return

        cursor.execute("""
            SELECT id
            FROM reviews
            WHERE order_id = ?
        """, (order_id,))

        existing_review = cursor.fetchone()

        if existing_review:

            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "⭐ شما قبلاً برای این سفارش "
                    "امتیاز ثبت کرده‌اید."
                )
            )

            return

        rating_comments[
            user_id
        ] = {
            "order_id": order_id,
            "product": product_code,
            "rating": rating
        }

        await bot.send_message(
            chat_id=chat_id,
            text=(
                f"⭐ امتیاز {rating} از 5 ثبت شد.\n\n"

                "اگر مایل هستید، نظر خود درباره محصول را "
                "در پیام بعدی بنویسید.\n\n"

                "اگر نمی‌خواهید نظری ثبت کنید، عبارت "
                "«بدون نظر» را ارسال کنید."
            )
        )

        return


# =========================================================
# ثبت نظر
# =========================================================

@bot.on_message()
async def handle_review_message(
    bot_instance: Robot,
    message: Message
):

    user_id = get_user_id(
        message
    )

    chat_id = get_chat_id(
        message
    )

    text = (
        getattr(
            message,
            "text",
            ""
        )
        or ""
    )

    if user_id not in rating_comments:
        return

    data = rating_comments[
        user_id
    ]

    order_id = data[
        "order_id"
    ]

    product_code = data[
        "product"
    ]

    rating = data[
        "rating"
    ]

    comment = text.strip()

    if comment == "بدون نظر":
        comment = ""

    cursor.execute("""
        SELECT id
        FROM reviews
        WHERE order_id = ?
    """, (order_id,))

    existing = cursor.fetchone()

    if existing:

        rating_comments.pop(
            user_id,
            None
        )

        return

    cursor.execute("""
        INSERT INTO reviews
        (
            order_id,
            user_id,
            product,
            rating,
            comment,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        order_id,
        user_id,
        product_code,
        rating,
        comment,
        now_string()
    ))

    conn.commit()

    rating_comments.pop(
        user_id,
        None
    )

    await bot.send_message(
        chat_id=chat_id,
        text=(
            "🙏 ممنون از ثبت نظر شما.\n\n"

            f"⭐ امتیاز شما: {rating} از 5\n"

            + (
                f"💬 نظر شما: {comment}\n\n"
                if comment
                else "\n"
            )

            + "نظر شما به بهبود فروشگاه کمک می‌کند."
        ),
        chat_keypad=main_menu(),
        chat_keypad_type="New"
    )


# =========================================================
# اجرای ربات
# =========================================================

if __name__ == "__main__":

    print(
        "BOT STARTING..."
    )

    print(
        "Connecting to the server..."
    )

    bot.run()
