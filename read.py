import argparse
import os
from docx import Document


def extract_images_from_docx(docx_file, target_dir):
    doc = Document(docx_file)
    images = []

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image = rel.target_part.blob
            images.append(image)

    # 保存图片到目标目录
    for i, image in enumerate(images, start=1):
        image_path = os.path.join(target_dir, f"image{i}.png")
        with open(image_path, "wb") as f:
            f.write(image)

    return len(images)


if __name__ == "__main__":
    # 创建命令行解析器
    parser = argparse.ArgumentParser(description="Extract images from Word document")
    parser.add_argument("-d", "--doc", help="Path to the Word document")
    parser.add_argument("-t", "--target", help="Target directory to save the images", default=".")
    args = parser.parse_args()

    if not os.path.exists(args.target):
        os.mkdir(args.target)

    # 提取图片并保存到目标目录
    num_images = extract_images_from_docx(args.doc, args.target)

    print(f"总共找到 {num_images} 张图片，已保存到目录：{args.target}")
