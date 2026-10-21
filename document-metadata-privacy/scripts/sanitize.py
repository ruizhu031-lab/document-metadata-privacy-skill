"""Local, copy-only document property cleanup; never a proof of authorship."""
import argparse
from io import BytesIO
import json
from pathlib import Path
import re
import sys
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

OOXML = {'.docx', '.pptx', '.xlsx'}
APP_FIELDS = {'Application', 'AppVersion', 'Company', 'Manager', 'Template', 'TotalTime'}
MARKERS = re.compile(rb'codex|openai|chatgpt|python-docx|python-pptx|pptxgenjs|reportlab|pypdf|openpyxl|wps|kingsoft', re.I)
MAX_BYTES = 512 * 1024 * 1024


def xml(data):
    from lxml import etree
    if b'<!DOCTYPE' in data or b'<!ENTITY' in data:
        raise ValueError('DTD/entity declarations are unsupported')
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def local(tag):
    return tag.rsplit('}', 1)[-1] if isinstance(tag, str) else ''


def package(path):
    with ZipFile(path) as z:
        infos = z.infolist()
        if len({i.filename for i in infos}) != len(infos):
            raise ValueError('Duplicate ZIP member names')
        if sum(i.file_size for i in infos) > MAX_BYTES or len(infos) > 20000:
            raise ValueError('Package exceeds inspection size limit')
        if any(i.flag_bits & 1 for i in infos):
            raise ValueError('Encrypted package is unsupported')
        parts = {i.filename: z.read(i) for i in infos}
        comment = bool(z.comment or any(i.comment or i.extra for i in infos))
    if '[Content_Types].xml' not in parts or '_rels/.rels' not in parts:
        raise ValueError('Not a recognized OOXML package')
    expected = {'.docx': 'word/document.xml', '.pptx': 'ppt/presentation.xml', '.xlsx': 'xl/workbook.xml'}
    if expected[path.suffix.lower()] not in parts:
        raise ValueError('Package content does not match its extension')
    return parts, comment


def property_parts(parts):
    kinds = {'docProps/core.xml': 'core', 'docProps/app.xml': 'app', 'docProps/custom.xml': 'custom'}
    for item in xml(parts['[Content_Types].xml']):
        ct = item.get('ContentType', '')
        for token, kind in (('core-properties', 'core'), ('extended-properties', 'app'), ('custom-properties', 'custom')):
            if token in ct:
                kinds[item.get('PartName', '').lstrip('/')] = kind
    return {name: kind for name, kind in kinds.items() if name in parts}


def inspect_office(path):
    parts, comments = package(path)
    kinds = property_parts(parts)
    fields, flags, markers = [], set(), []
    if comments:
        fields.append('ZIP:comments_or_extra_fields')
    for name, kind in kinds.items():
        root = xml(parts[name])
        for child in root:
            label = local(child.tag)
            if kind != 'app' or label in APP_FIELDS:
                if ''.join(child.itertext()).strip() or child.attrib:
                    fields.append(name + ':' + label)
        if kind == 'custom' and len(root):
            flags.add('custom_properties_may_feed_fields')
    for name, data in parts.items():
        lower = name.lower()
        if lower.startswith('_xmlsignatures/') or lower.endswith('vbaprojectsignature.bin'):
            flags.add('digital_signature')
        if 'comment' in lower or 'persons/' in lower or lower.endswith('person.xml'):
            flags.add('comments_or_people')
        if 'notesslides/' in lower:
            flags.add('speaker_notes')
        if 'embeddings/' in lower:
            flags.add('embedded_files')
        if '/media/' in lower:
            flags.add('embedded_media_metadata_not_inspected')
        if lower.startswith('customxml/'):
            flags.add('custom_xml')
        if 'thumbnail' in lower:
            flags.add('thumbnail_preview')
        if lower.endswith(('.xml', '.rels')):
            root = xml(data)
            if MARKERS.search(data):
                markers.append(name)
            for node in root.iter():
                tag = local(node.tag)
                attrs = {local(k): v for k, v in node.attrib.items()}
                if tag in {'ins', 'del', 'moveFrom', 'moveTo'} or any(k in attrs for k in ('author', 'initials', 'userId', 'providerId')):
                    flags.add('revision_or_identity_attributes')
                if tag == 'Relationship' and attrs.get('TargetMode') == 'External':
                    flags.add('external_links_or_template')
                if tag == 'sheet' and attrs.get('state') in {'hidden', 'veryHidden'}:
                    flags.add('hidden_sheets')
                if tag in {'vanish', 'webHidden'} or (tag == 'sld' and attrs.get('show') == '0'):
                    flags.add('hidden_content')
                if tag == 'HyperlinkBase' and node.text:
                    flags.add('hyperlink_base_preserved_for_link_integrity')
            if re.search(rb'DOCPROPERTY|DOCVARIABLE', data, re.I):
                flags.add('document_property_fields')
    return {'format': path.suffix.lower(), 'metadata_fields': sorted(fields), 'review_flags': sorted(flags), 'marker_parts': sorted(markers)}


def office_bytes(path):
    from lxml import etree
    parts, _ = package(path)
    before = inspect_office(path)
    if 'digital_signature' in before['review_flags']:
        raise ValueError('Signed document: metadata edits would invalidate the signature')
    for name, kind in property_parts(parts).items():
        root = xml(parts[name])
        if kind != 'app':
            # Rebuild only the property root, dropping property values and XML comments.
            root = etree.Element(root.tag, nsmap=root.nsmap)
        else:
            for child in list(root):
                if local(child.tag) in APP_FIELDS or not isinstance(child.tag, str):
                    root.remove(child)
        parts[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w', ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = ZipInfo(name)  # Neutral ZIP time, no source comments or extended attributes.
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    return buffer.getvalue(), before


def pdf_nodes(root):
    from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject
    seen, pending = set(), [root]
    while pending:
        node = pending.pop()
        if isinstance(node, IndirectObject):
            node = node.get_object()
        if id(node) in seen:
            continue
        seen.add(id(node))
        if isinstance(node, DictionaryObject):
            yield node
            pending.extend(node.values())
        elif isinstance(node, ArrayObject):
            pending.extend(node)


def read_pdf(path):
    from pypdf import PdfReader
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise ValueError('Encrypted PDF is unsupported; use an authorized decrypted copy')
    return reader


def inspect_pdf(path):
    reader = read_pdf(path)
    fields = sorted(str(k) for k in (reader.metadata or {}))
    flags, metadata_count = set(), 0
    for node in pdf_nodes(reader.trailer['/Root']):
        if '/Metadata' in node:
            metadata_count += 1
            raw = node['/Metadata'].get_data()
            if b'pdfaid' in raw.lower() or b'pdfuaid' in raw.lower():
                flags.add('archival_or_accessibility_conformance')
        if '/ByteRange' in node or node.get('/FT') == '/Sig' or node.get('/Type') == '/Sig':
            flags.add('digital_signature')
        if '/Annots' in node:
            flags.add('annotations')
        if '/EmbeddedFiles' in node or '/EF' in node:
            flags.add('embedded_files')
        if '/AcroForm' in node:
            flags.add('forms')
        if node.get('/Subtype') == '/Image':
            flags.add('embedded_media_metadata_not_inspected')
    if metadata_count:
        fields.append('XMP_streams:' + str(metadata_count))
    return {'format': '.pdf', 'metadata_fields': fields, 'review_flags': sorted(flags), 'marker_parts': [], 'pages': len(reader.pages), 'scan_scope': 'Info and reachable Metadata streams; visible content and arbitrary PDF strings are not identity-scanned'}


def pdf_bytes(path):
    from pypdf import PdfWriter
    before = inspect_pdf(path)
    if {'digital_signature', 'archival_or_accessibility_conformance'} & set(before['review_flags']):
        raise ValueError('Signed or PDF/A/PDF/UA document requires a format-specific workflow')
    writer = PdfWriter(clone_from=read_pdf(path))
    writer.metadata = None
    for node in list(pdf_nodes(writer.root_object)):
        node.pop('/Metadata', None)
    # Full rewrite and orphan collection, not an incremental update retaining old objects.
    writer.compress_identical_objects(remove_identicals=False, remove_orphans=True)
    buffer = BytesIO()
    writer.write(buffer)
    return buffer.getvalue(), before


def inspect(path):
    if path.stat().st_size > MAX_BYTES:
        raise ValueError('File exceeds inspection size limit')
    if path.suffix.lower() in OOXML:
        return inspect_office(path)
    if path.suffix.lower() == '.pdf':
        return inspect_pdf(path)
    raise ValueError('Unsupported format. Convert legacy Office/WPS files using the originating app; never rename extensions')


def clean(source, output):
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError('Output must be a new file distinct from the source')
    if source.suffix.lower() != output.suffix.lower():
        raise ValueError('Output extension must match source')
    inspect(source)
    payload, before = office_bytes(source) if source.suffix.lower() in OOXML else pdf_bytes(source)
    created = False
    try:
        with output.open('xb') as stream:
            created = True
            stream.write(payload)
        remaining = inspect(output)
        if remaining['metadata_fields']:
            raise ValueError('Output property verification failed')
    except Exception:
        if created:
            output.unlink(missing_ok=True)
        raise
    return {'status': 'targeted_metadata_removed', 'output': str(output.resolve()), 'before': before, 'remaining': remaining, 'limits': 'Review flags and content remain. This is not an AI-origin, anonymity, forensic, or full hidden-data guarantee.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['inspect', 'clean'])
    parser.add_argument('source', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    try:
        if args.action == 'clean' and args.output is None:
            raise ValueError('clean requires --output')
        report = inspect(args.source) if args.action == 'inspect' else clean(args.source, args.output)
        print(json.dumps(report, ensure_ascii=True, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({'status': 'error', 'error': str(exc)}, ensure_ascii=True))
        return 2


if __name__ == '__main__':
    sys.exit(main())
