from zipfile import ZipFile, ZIP_DEFLATED
# copy

def appendXMLElement(xml_str:str, element_xml):
    insert_idx = xml_str.index('</')
    return xml_str[:insert_idx] + element_xml + xml_str[insert_idx:]
    
def injectTitus(filepath):
    with ZipFile(filepath, 'r') as zip:
        content_types = zip.read(r'[Content_Types].xml').decode()
        _rels = zip.read(r'_rels/.rels').decode()
        _rels_idx = _rels.count('Relationship')-1
        label_info = '<?xml version="1.0" encoding="utf-8" standalone="yes"?><clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"><clbl:label id="{cccd100a-077b-4351-b7ea-99b99562cb12}" enabled="1" method="Privileged" siteId="{f06fa858-824b-4a85-aacb-f372cfdc282e}" contentBits="0" removed="0" /></clbl:labelList>'
    
        excluded_files = [r'[Content_Types].xml', r'_rels/.rels']
    
        other_data = {}
        for item in zip.infolist():
            if item.filename not in excluded_files:
                # preserve compression & metadata
                data = zip.read(item.filename)
                other_data[item.filename] = data

    with ZipFile(filepath, 'w', compression=ZIP_DEFLATED) as zip:

        # add other files
        for file, data in other_data.items():
            zip.writestr(file, data)

        # update content types
        labelInfo_CTypeXML='<Override PartName="/docMetadata/LabelInfo.xml" ContentType="application/vnd.ms-office.classificationlabels+xml" />'
        if 'docMetadata/LabelInfo.xml' not in content_types:
            zip.writestr(r'[Content_Types].xml', appendXMLElement(content_types, labelInfo_CTypeXML))
        else:
            zip.writestr(r'[Content_Types].xml', content_types)

        # update relationship
        labelInfo_relsXML=f'<Relationship Id="rId{_rels_idx}"  Type="http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels"  Target="docMetadata/LabelInfo.xml" />'
        if 'docMetadata/LabelInfo.xml' not in _rels:
            zip.writestr(r'_rels/.rels', appendXMLElement(_rels, labelInfo_relsXML))
        else:
            zip.writestr(r'_rels/.rels', _rels)

        # insert label
        zip_filelist = [f.filename for f in zip.filelist]
        if r'docMetadata/LabelInfo.xml' not in zip_filelist:
            zip.writestr(r'docMetadata/LabelInfo.xml', label_info)

import argparse
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Inject Titus label metadata into XLSX"
    )
    parser.add_argument("filepath", help="Path to the XLSX file to modify")
    args = parser.parse_args()

    try:
        injectTitus(args.filepath)
        print(f"Successfully injected security label 'Public' into: {args.filepath}")
    except Exception as e:
        print(f"Failed to inject security label into {args.filepath}: {e}")
        raise
