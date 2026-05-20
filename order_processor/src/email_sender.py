"""メール送信（Excelレポートを添付）"""
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime


def send_report(
    output_path: str,
    config: dict,
    order_count: int,
    unknown_count: int,
):
    """レポートExcelを田村さん・脇阪さんにメール送信する"""
    subject, body = _build_message(order_count, unknown_count)

    msg = MIMEMultipart()
    msg["From"]    = config["sender_address"]
    msg["To"]      = ", ".join(config["recipients"].values())
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Excelを添付
    with open(output_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{os.path.basename(output_path)}"',
    )
    msg.attach(part)

    with smtplib.SMTP(config["smtp_server"], config["smtp_port"]) as server:
        server.starttls()
        server.login(config["sender_address"], config["sender_password"])
        server.sendmail(
            config["sender_address"],
            list(config["recipients"].values()),
            msg.as_string(),
        )

    return list(config["recipients"].keys())


def _build_message(order_count: int, unknown_count: int) -> tuple[str, str]:
    now = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    subject = f"【注文解析レポート】{datetime.now().strftime('%Y/%m/%d')} 生成分"

    lines = [
        f"お世話になっております。",
        f"",
        f"{now} に注文解析レポートを生成しました。",
        f"",
        f"■ 集計結果",
        f"　対象注文件数: {order_count}件",
    ]
    if unknown_count > 0:
        lines.append(f"　構成未登録品番: {unknown_count}件（Excelの集計シートをご確認ください）")
    else:
        lines.append(f"　構成未登録品番: なし")

    lines += [
        f"",
        f"詳細は添付のExcelファイルをご確認ください。",
        f"",
        f"※このメールは自動送信されています。",
    ]

    return subject, "\n".join(lines)
