from excelUtil import embed_image


def main():
    embed_image("old.xlsx", "new.xlsx", "Sheet1", "图片")


if __name__ == '__main__':
    main()
