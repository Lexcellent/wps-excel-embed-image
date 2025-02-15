import os
import shutil
import zipfile
from uuid import uuid4

import lxml
import pandas as pd
from PIL import Image
from loguru import logger
from lxml import etree

cellimages_rels_template_content = """
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"""

cellimages_template_content = """
<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">

</etc:cellImages>
"""


def unzip_file(zip_path: str, extract_to: str):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    logger.debug(f"Extracted {zip_path} to {extract_to}")


def zip_file(file_or_dir_path: str, zip_path: str):
    # Create a ZipFile object in write mode
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        if os.path.isfile(file_or_dir_path):
            # If the path is a file, write it to the zip file
            zipf.write(file_or_dir_path, os.path.basename(file_or_dir_path))
        elif os.path.isdir(file_or_dir_path):
            # If the path is a directory, walk through the directory
            for root, dirs, files in os.walk(file_or_dir_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, file_or_dir_path)
                    zipf.write(file_path, arcname)
    logger.debug(f"Compressed {file_or_dir_path} to {zip_path}")


def get_image_dimensions(image_path: str):
    with Image.open(image_path) as img:
        width, height = img.size
        logger.debug(f"Image dimensions: width={width}, height={height}")
        return width, height


def copy_image_to_excel_dir(image_path: str, excel_unzip_dir: str):
    media_dir = os.path.join(excel_unzip_dir, 'xl', 'media')
    if not os.path.exists(media_dir):
        os.makedirs(media_dir)

    image_filename = os.path.basename(image_path)
    destination_path = os.path.join(media_dir, image_filename)
    shutil.copy(image_path, destination_path)
    logger.debug(f"Copied {image_path} to {destination_path}")
    return image_filename, destination_path


def add_new_node_cell_images(image_path: str, excel_unzip_dir: str):
    # 指定xml是否存在 excel中未插入过图片则不存在 需要补充进去
    xmlPath = os.path.join(excel_unzip_dir, 'xl', 'cellimages.xml')
    if not os.path.exists(xmlPath):
        # 模板内容写入
        with open(xmlPath, 'w', encoding='utf-8') as f:
            f.write(cellimages_template_content)

    uuid = str(uuid4()).replace('-', '').upper()
    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    tree = etree.parse(xmlPath, parser=parser)
    root = tree.getroot()
    # 提取命名空间
    namespaces = {prefix: uri for prefix, uri in root.nsmap.items() if prefix is not None}
    logger.debug(f"Extracted namespaces: {namespaces}")

    count = len(root.findall('.//xdr:pic', namespaces=namespaces))
    ID = f"ID_{uuid}"
    RID = f"rId{count + 1}"
    width, height = get_image_dimensions(image_path)
    xml_string = f"""
    <etc:cellImage>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id="{count + 1}" name="{ID}" descr="{os.path.basename(image_path)}"/>
                <xdr:cNvPicPr/>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip r:embed="{RID}"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="{int(width * 1000)}" cy="{int(height * 1000)}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
    </etc:cellImage>
    """
    new_cell_image = etree.fromstring(xml_string, parser=parser)
    root.append(new_cell_image)
    tree.write(xmlPath, pretty_print=True, xml_declaration=True, encoding='UTF-8')

    return ID, RID


def add_new_node_cell_images_rels(image_file_name: str, RID: str, excel_unzip_dir: str):
    # 指定xml是否存在 excel中未插入过图片则不存在 需要补充进去
    xmlPath = os.path.join(excel_unzip_dir, 'xl', '_rels', 'cellimages.xml.rels')
    if not os.path.exists(xmlPath):
        # 模板内容写入
        with open(xmlPath, 'w', encoding='utf-8') as f:
            f.write(cellimages_rels_template_content)

    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    xml_string = f"""
       <Relationship Id="{RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{image_file_name}"/>
    """
    new_cell_image = etree.fromstring(xml_string, parser=parser)
    tree = etree.parse(xmlPath, parser=parser)
    tree.getroot().append(new_cell_image)
    tree.write(xmlPath, pretty_print=True, xml_declaration=True, encoding='UTF-8')


def add_new_node_content_types(excel_unzip_dir: str):
    # 解析 XML 文件
    xmlPath = os.path.join(excel_unzip_dir, '[Content_Types].xml')
    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    tree = etree.parse(xmlPath, parser=parser)
    root = tree.getroot()
    if root.find('{http://schemas.openxmlformats.org/package/2006/content-types}Default[@Extension="JPG"]') is None:
        xml_string = f"""
                 <Default Extension="JPG" ContentType="image/.jpg"/>
            """
        new_cell_image = etree.fromstring(xml_string, parser=parser)
        tree.getroot().append(new_cell_image)
    if root.find('{http://schemas.openxmlformats.org/package/2006/content-types}Default[@Extension="jpeg"]') is None:
        xml_string = f"""
                 <Default Extension="jpeg" ContentType="image/jpeg"/>
            """
        new_cell_image = etree.fromstring(xml_string, parser=parser)
        tree.getroot().append(new_cell_image)
    if root.find('{http://schemas.openxmlformats.org/package/2006/content-types}Override[@PartName="/xl/cellimages.xml"]') is None:
        xml_string = f"""
                 <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>
            """
        new_cell_image = etree.fromstring(xml_string, parser=parser)
        tree.getroot().append(new_cell_image)

    tree.write(xmlPath, pretty_print=True, xml_declaration=True, encoding='UTF-8')


def add_new_node_workbook(excel_unzip_dir: str):
    # 解析 XML 文件
    xmlPath = os.path.join(excel_unzip_dir, 'xl', '_rels', 'workbook.xml.rels')
    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    tree = etree.parse(xmlPath, parser=parser)
    root = tree.getroot()
    if root.find('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship[@Target="cellimages.xml"]') is None:
        xml_string = f"""
                 <Relationship Id="rId100" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>
            """
        new_cell_image = etree.fromstring(xml_string, parser=parser)
        tree.getroot().append(new_cell_image)

    tree.write(xmlPath, pretty_print=True, xml_declaration=True, encoding='UTF-8')


def add_new_node(image_path: str, unzip_dir_path: str):
    add_new_node_content_types(unzip_dir_path)
    add_new_node_workbook(unzip_dir_path)

    image_name, image_path = copy_image_to_excel_dir(image_path, unzip_dir_path)
    ID, RID = add_new_node_cell_images(image_path, unzip_dir_path)
    add_new_node_cell_images_rels(image_name, RID, unzip_dir_path)
    return ID


def add_sheet_data(unzip_file_path, sheet_name, ID, cell_index, row_index):
    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    xmlPath = os.path.join(unzip_file_path, 'xl', 'worksheets', f'{sheet_name}.xml')
    tree = etree.parse(xmlPath, parser=parser)
    root = tree.getroot()
    sheetDataTag = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData')
    logger.debug(sheetDataTag)
    rowTag = sheetDataTag.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')[row_index]
    cellTag: lxml.etree._Element = rowTag.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')[cell_index]
    r = cellTag.attrib['r']
    cellTag.clear()
    xml_string = f"""
                <f>_xlfn.DISPIMG(&quot;{ID}&quot;,1)</f>
    """
    new_tag = etree.fromstring(xml_string, parser=parser)
    cellTag.append(new_tag)
    xml_string = f"""
               <v>=DISPIMG(&quot;{ID}&quot;,1)</v>
    """
    new_tag = etree.fromstring(xml_string, parser=parser)
    cellTag.append(new_tag)
    cellTag.attrib['r'] = r
    cellTag.attrib['t'] = "str"
    # cellTag.remove('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    tree.write(xmlPath, pretty_print=True, xml_declaration=True, encoding='UTF-8')


def embed_image(excel_path: str, new_excel_path: str, sheet_name: str, head_name: str):
    # 解压excel
    unzip_dir_path = f"{excel_path}excelUnZipDir"
    if os.path.exists(unzip_dir_path):
        shutil.rmtree(unzip_dir_path)
    unzip_file(excel_path, unzip_dir_path)
    # 读取Excel文件
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    column_index = df.columns.get_loc(head_name)
    logger.debug(f"Column '{head_name}' is at index: {column_index}")
    for index in range(len(df.get(head_name))):
        pic_local_path = str(df.get(head_name)[index])
        logger.debug(f"Picture row: {pic_local_path}")
        if pic_local_path is None or not os.path.exists(pic_local_path):
            logger.debug(f"图片文件 {pic_local_path} 不存在")
            continue
        ID = add_new_node(pic_local_path, unzip_dir_path)
        add_sheet_data(unzip_dir_path, sheet_name, ID, column_index, index + 1)
    zip_file(unzip_dir_path, new_excel_path)

    if os.path.exists(unzip_dir_path):
        shutil.rmtree(unzip_dir_path)
