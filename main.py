import os
import shutil
import zipfile
from uuid import uuid4

import lxml
import pandas as pd
from PIL import Image
from loguru import logger
from lxml import etree
from pandas import DataFrame


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
    # 解析 XML 文件
    xmlPath = os.path.join(excel_unzip_dir, 'xl', 'cellimages.xml')
    if not os.path.exists(xmlPath):
        # 拷贝模板文件到这个路径
        shutil.copy("cellimages_template.xml", xmlPath)

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
    # 解析 XML 文件
    xmlPath = os.path.join(excel_unzip_dir, 'xl', '_rels', 'cellimages.xml.rels')
    if not os.path.exists(xmlPath):
        # 拷贝模板文件到这个路径
        shutil.copy("cellimages.xml_template.rels", xmlPath)

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


def add_new_node(image_path: str, unzip_file_path: str):
    add_new_node_content_types(unzip_file_path)
    add_new_node_workbook(unzip_file_path)

    image_name, image_path = copy_image_to_excel_dir(image_path, unzip_file_path)
    ID, RID = add_new_node_cell_images(image_path, unzip_file_path)
    add_new_node_cell_images_rels(image_name, RID, unzip_file_path)
    return ID


def get_cell_word(unzip_file_path: str, sheet_name: str, cell_name: str):
    # 创建解析器时关闭 DTD 和校验
    parser = etree.XMLParser(
        load_dtd=False,  # 不加载 DTD
        no_network=True,  # 禁止网络请求（防止自动下载 DTD）
        dtd_validation=False,  # 关闭 DTD 校验
        attribute_defaults=False,
        recover=True
    )
    tree = etree.parse(os.path.join(unzip_file_path, 'xl', 'sharedStrings.xml'), parser=parser)
    root = tree.getroot()
    sis = root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
    index = 0
    for si in sis:
        if si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text == cell_name:
            return index
        index += 1
    return index


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


def embed_image(excel_path: str, new_excel_path: str, sheet_name: str, cell_name: str):
    # 解压excel
    unzip_file_path = "excelUnZipDir"
    if os.path.exists(unzip_file_path):
        shutil.rmtree(unzip_file_path)
    unzip_file(excel_path, unzip_file_path)

    cell_index = get_cell_word(unzip_file_path, sheet_name, cell_name)
    logger.debug(f"cell_index: {cell_index}")
    df: DataFrame = pd.read_excel(excel_path, sheet_name=sheet_name)
    for index in range(len(df.get(cell_name))):
        picRow = df.get(cell_name)[index]
        ID = add_new_node(picRow, unzip_file_path)
        add_sheet_data(unzip_file_path, sheet_name, ID, cell_index, index + 1)
    zip_file(unzip_file_path, new_excel_path)

    if os.path.exists(unzip_file_path):
        shutil.rmtree(unzip_file_path)


def main():
    embed_image("old.xlsx", "new.xlsx", "Sheet1", "pic")


if __name__ == '__main__':
    main()
