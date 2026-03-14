import json
import os
os.environ["DISABLE_MODEL_SOURCE_CHECK"] = "True"
os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"
import shutil

def recognize_image_and_export_markdown(image_path, output_path):
    """
    解析图像,识别文字并导出成markdown文档(可能有部分识别错误)

    Args:
        image_path (str): 图像文件路径
        output_path (str): 导出的markdown文档的路径

    Returns:
        str: 状态
    """
    try:
        from paddleocr import PaddleOCRVL

        # 英伟达 GPU
        pipeline = PaddleOCRVL()

        output = pipeline.predict(image_path)
        for res in output:
            # res.print()  ## 打印预测的结构化输出
            # res.save_to_json(save_path="output")  ## 保存当前图像的结构化json结果
            res.save_to_markdown(save_path="output")  ## 保存当前图像的markdown格式的结果
        name = image_path.split("\\")[-1].split("/")[-1].split(".")[-2]
        file =  os.path.abspath("./output/%s.md" % (name))
        shutil.copy(file, output_path)
        return "识别成功,已导出至%s"%(output_path)
    except Exception as e:
        return f"错误: {str(e)}"


if __name__ == "__main__":
    recognize_image_and_export_markdown("C:\\Users\\22974\\Desktop\\作者唐希鲲_screen_capture.png","./b.md")
    # print(
    #     recognize_image_and_export_markdown(
    #         "C:/Users/22974/Desktop/单词表.jpg", "C:/Users/22974/Desktop/单词表内容.txt"
    #     )
    # )
    pass
