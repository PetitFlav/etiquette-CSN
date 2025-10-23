from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from email.message import EmailMessage
import smtplib
from pathlib import Path
from typing import Callable, ContextManager, Iterable, Mapping
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from jinja2 import Template

from .config import ATTESTATION_TEMPLATE_PATH


DEFAULT_SUBJECT = "Attestation de paiement - {{ prenom }} {{ nom }}"
DEFAULT_BODY = (
    "Bonjour {{ prenom }},\n\n"
    "Veuillez trouver ci-joint votre attestation de paiement.\n"
    "Montant réglé : {{ montant }}.\n\n"
    "Cordialement."
)


@dataclass
class AttestationData:
    nom: str
    prenom: str
    email: str
    montant: str
    expire: str = ""
    date_de_naissance: str = ""
    generated_at: datetime = field(default_factory=datetime.utcnow)

    def context(self) -> dict[str, object]:
        return {
            "nom": self.nom,
            "prenom": self.prenom,
            "email": self.email,
            "montant": self.montant,
            "expire": self.expire,
            "date_de_naissance": self.date_de_naissance,
            "created_at": self.generated_at,
            "created_at_display": self.generated_at.strftime("%d/%m/%Y"),
        }


@dataclass
class SMTPSettings:
    host: str
    port: int
    sender: str
    username: str = ""
    password: str = ""
    use_tls: bool = True
    use_ssl: bool = False
    timeout: float = 30.0
    subject_template: str = DEFAULT_SUBJECT
    body_template: str = DEFAULT_BODY

    @property
    def is_configured(self) -> bool:
        return bool(self.host and self.sender)

    def render_subject(self, data: AttestationData) -> str:
        template = Template(self.subject_template)
        return template.render(**data.context())

    def render_body(self, data: AttestationData) -> str:
        template = Template(self.body_template)
        return template.render(**data.context())


def _parse_int(value: str | None, default: int) -> int:
    if not value:
        return default
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _parse_bool(value: str | None, *, default: bool) -> bool:
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on", "oui"}:
        return True
    if normalized in {"0", "false", "no", "n", "off", "non"}:
        return False
    return default


def load_attestation_settings(config: Mapping[str, str]) -> SMTPSettings:
    host = (config.get("smtp_host", "") or "").strip()
    sender = (config.get("smtp_sender", "") or "").strip()
    port = _parse_int(config.get("smtp_port"), default=0)
    username = (config.get("smtp_user", "") or "").strip()
    password = config.get("smtp_password", "") or ""
    timeout = float(_parse_int(config.get("smtp_timeout"), default=30))
    use_tls = _parse_bool(config.get("smtp_use_tls"), default=True)
    use_ssl = _parse_bool(config.get("smtp_use_ssl"), default=False)
    subject_template = config.get("attestation_subject") or DEFAULT_SUBJECT
    body_template = config.get("attestation_body") or DEFAULT_BODY

    if use_ssl and port == 0:
        port = 465
    elif port == 0:
        port = 587 if use_tls else 25

    return SMTPSettings(
        host=host,
        port=port,
        sender=sender,
        username=username,
        password=password,
        use_tls=use_tls,
        use_ssl=use_ssl,
        timeout=timeout,
        subject_template=subject_template,
        body_template=body_template,
    )


def _sanitize_filename(value: str) -> str:
    sanitized = "".join(ch for ch in value if ch.isalnum() or ch in {"_", "-"})
    return sanitized or "attestation"


def _escape_pdf_text(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return escaped


def _docx_template_lines(data: AttestationData, template_path: Path | None) -> list[str]:
    """Return attestation lines rendered from the DOCX template.

    When the template is missing or unparsable, a textual fallback is used.
    """

    try:
        path = template_path or ATTESTATION_TEMPLATE_PATH
        with ZipFile(path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8")
    except FileNotFoundError:
        return _fallback_template_lines(data)
    except KeyError:
        return _fallback_template_lines(data)

    rendered = _apply_template_replacements(xml, data)
    try:
        return _extract_paragraphs(rendered)
    except ET.ParseError:
        return _fallback_template_lines(data)


def _fallback_template_lines(data: AttestationData) -> list[str]:
    montant_text = _normalize_montant_value(data.montant)
    return [
        "Attestation de paiement",
        "",
        f"Nous certifions que {data.prenom} {data.nom}",
        f"a réglé pour la somme de {montant_text} au titre de son inscription.",
        "",
        f"Fait à Nantes, le {data.generated_at.strftime('%d/%m/%Y')}",
        "",
        "Signature",
    ]


def _normalize_montant_value(raw: str) -> str:
    montant = (raw or "").strip()
    if not montant:
        return ""
    if montant.endswith("€"):
        return montant
    return f"{montant} €"


def _apply_template_replacements(xml: str, data: AttestationData) -> str:
    montant_text = _normalize_montant_value(data.montant)
    full_name = f"{data.prenom} {data.nom}".strip()
    date_text = data.generated_at.strftime("%d/%m/%Y")

    replacements = [
        ("Mr, Mme, Melle .......", f"Mr, Mme, Melle {full_name}" if full_name else "Mr, Mme, Melle"),
        (
            "pour la somme de .....€",
            f"pour la somme de {montant_text or '_____ €'}",
        ),
        ("Fait à Nantes, le ", f"Fait à Nantes, le {date_text}"),
    ]

    rendered = xml
    for needle, value in replacements:
        rendered = rendered.replace(needle, value)
    return rendered


def _extract_paragraphs(xml: str) -> list[str]:
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    root = ET.fromstring(xml)
    body = root.find("w:body", namespaces)
    if body is None:
        return []

    lines: list[str] = []
    for paragraph in body.findall("w:p", namespaces):
        texts = [node.text for node in paragraph.findall('.//w:t', namespaces) if node.text]
        lines.append("".join(texts))

    return lines


def _merge_additional_metadata(lines: list[str], data: AttestationData) -> list[str]:
    enriched = list(lines)
    extra_lines: list[str] = []
    if data.expire:
        extra_lines.append(f"Expiration : {data.expire}")
    if data.date_de_naissance:
        extra_lines.append(f"Date de naissance : {data.date_de_naissance}")
    if extra_lines:
        if enriched and enriched[-1].strip():
            enriched.append("")
        enriched.extend(extra_lines)
    return enriched


def _build_pdf_stream_lines(data: AttestationData) -> Iterable[str]:
    lines = _docx_template_lines(data, ATTESTATION_TEMPLATE_PATH)

    if not lines:
        lines = _fallback_template_lines(data)

    lines = _merge_additional_metadata(lines, data)

    y = 780
    first_text_idx = next((idx for idx, value in enumerate(lines) if value.strip()), 0)
    for idx, line in enumerate(lines):
        font_size = 18 if idx == first_text_idx else 12
        yield "BT"
        yield f"/F1 {font_size} Tf"
        yield f"72 {y} Td"
        yield f"({_escape_pdf_text(line)}) Tj"
        yield "ET"
        y -= 28 if idx == first_text_idx else 20


def build_attestation_pdf_bytes(data: AttestationData) -> bytes:
    content_lines = list(_build_pdf_stream_lines(data))
    content_stream = "\n".join(content_lines).encode("cp1252", "replace")

    objects: list[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    objects.append(b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    objects.append(
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
        b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"
    )
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    content_obj = (
        f"<< /Length {len(content_stream)} >>\nstream\n".encode("ascii")
        + content_stream
        + b"\nendstream"
    )
    objects.append(content_obj)

    pdf = bytearray()
    pdf.extend(b"%PDF-1.4\n")
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{index} 0 obj\n".encode("ascii"))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")

    xref_offset = len(pdf)
    count = len(objects) + 1
    pdf.extend(f"xref\n0 {count}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))

    pdf.extend(
        (
            "trailer\n"
            f"<< /Size {count} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n"
            "%%EOF"
        ).encode("ascii")
    )
    return bytes(pdf)


def generate_attestation_pdf(directory: Path, data: AttestationData) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    timestamp = data.generated_at.strftime("%Y%m%d-%H%M%S")
    nom_part = _sanitize_filename(data.nom.upper())
    prenom_part = _sanitize_filename(data.prenom.upper())
    filename = f"attestation_{prenom_part}_{nom_part}_{timestamp}.pdf"
    pdf_bytes = build_attestation_pdf_bytes(data)
    target = directory / filename
    target.write_bytes(pdf_bytes)
    return target


def build_email_message(settings: SMTPSettings, data: AttestationData, pdf_path: Path) -> EmailMessage:
    message = EmailMessage()
    message["Subject"] = settings.render_subject(data)
    message["From"] = settings.sender
    message["To"] = data.email
    message.set_content(settings.render_body(data))

    payload = pdf_path.read_bytes()
    message.add_attachment(
        payload,
        maintype="application",
        subtype="pdf",
        filename=pdf_path.name,
    )
    return message


@contextmanager
def _smtp_connection(settings: SMTPSettings):
    if not settings.host:
        raise RuntimeError("Serveur SMTP non configuré")

    port = settings.port or (465 if settings.use_ssl else 587 if settings.use_tls else 25)
    smtp_cls: Callable[..., smtplib.SMTP]
    if settings.use_ssl:
        smtp_cls = smtplib.SMTP_SSL
    else:
        smtp_cls = smtplib.SMTP

    try:
        smtp = smtp_cls(settings.host, port, timeout=settings.timeout)
    except OSError as exc:  # pragma: no cover - depends on environment
        raise RuntimeError(f"Connexion SMTP impossible: {exc}") from exc

    try:
        if not settings.use_ssl and settings.use_tls:
            smtp.starttls()
        if settings.username:
            smtp.login(settings.username, settings.password)
        yield smtp
    finally:  # pragma: no cover - best effort shutdown
        try:
            smtp.quit()
        except Exception:
            smtp.close()


def send_attestation_email(
    settings: SMTPSettings,
    data: AttestationData,
    pdf_path: Path,
    *,
    smtp_factory: Callable[[SMTPSettings], ContextManager[smtplib.SMTP]] | None = None,
) -> EmailMessage:
    message = build_email_message(settings, data, pdf_path)
    factory = smtp_factory or _smtp_context_wrapper

    try:
        with factory(settings) as smtp:
            smtp.send_message(message)
    except smtplib.SMTPException as exc:  # pragma: no cover - depends on SMTP backend
        raise RuntimeError(f"Envoi email impossible: {exc}") from exc

    return message


def _smtp_context_wrapper(settings: SMTPSettings):
    return _smtp_connection(settings)


__all__ = [
    "AttestationData",
    "SMTPSettings",
    "build_attestation_pdf_bytes",
    "build_email_message",
    "generate_attestation_pdf",
    "load_attestation_settings",
    "send_attestation_email",
]
