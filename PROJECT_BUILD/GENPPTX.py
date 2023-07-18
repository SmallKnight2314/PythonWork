from collections.abc import Container
from pptx import Presentation
from gen_pptx_methods.GenTitleSlide import GenTitleSlide
from gen_pptx_methods.GenTextSlide import GenTextSlide
from gen_pptx_methods.GenListSlide import GenListSlide
from gen_pptx_methods.GenPictureSlide import GenPictureSlide
from gen_pptx_methods.GenPlotSlide import GenPlotSlide
from UnpackJSON import UnpackJSON
import logging
import logging.config
logging.config.fileConfig("../debug/logging.conf")
logger = logging.getLogger("GENPPTX")

# extract JSON file contents
json_path = input("enter JSON file path to generate PPTX file : ")
pJSON = UnpackJSON(json_path)
extract = pJSON.extract()
pages = extract["pages"]
print(pages)

name = input("Save PPTX file, please provide a name without the .pptx : ")
filename = name + ".pptx"
logger.debug(f"filename is {filename}")
file_path = f"./output/{filename}"
logger.debug(file_path)
powerpoint = Presentation()
powerpoint.save(file_path)
logger.debug(f"{filename} has been saved at : {file_path}")

# begin the for loop with a switch/case
for page in pages:
    match page["page_type"]:
        case "TITLE_SLIDE":
            logger.debug("this is where Gen_Title_Slide() goes")
            title = page["page_content"]["title"]
            subtitle = page["page_content"]["subtitle"]
            pageNum = page["page#"]
            TITLE_SLIDE = GenTitleSlide(title, subtitle,pageNum,file_path)
            TITLE_SLIDE.generate()
            logger.debug("method complete!")

        case "TEXT_SLIDE":
            logger.debug("this is where Gen_Text_Slide() goes")
            title = page["page_content"]["title"]
            text = page["page_content"]["paragraphs"]
            pageNum = page["page#"]
            TEXT_SLIDE = GenTextSlide(title,text, pageNum, file_path)
            TEXT_SLIDE.generate()
            logger.debug("method complete!")

        case "LIST_SLIDE":
            logger.debug("this is where Gen_List_Slide() goes")
            title = page["page_content"]["title"]
            listData = page["page_content"]["lists"]
            pageNum = page["page#"]
            LIST_SLIDE = GenListSlide(title, listData, pageNum, file_path)
            LIST_SLIDE.generate()
            logger.debug("method complete!")

        case "PICTURE_SLIDE":
            logger.debug("this is where Gen_Picture_Slide() goes")
            title = page["page_content"]["title"]
            picture = page["page_content"]["image"]
            pageNum = page["page#"]
            caption = page["page_content"]["caption"]
            PICTURE_SLIDE = GenPictureSlide(title, picture, pageNum, caption, file_path)
            PICTURE_SLIDE.generate()
            logger.debug("method complete!")

        case "PLOT_SLIDE":
            logger.debug("this is where Gen_Plot_Slide() goes")
            title = page["page_content"]["title"]
            plotData = page["page_content"]["data"]
            pageNum = page["page#"]
            PICTURE_SLIDE = GenPlotSlide(title, plotData, pageNum, file_path)
            PICTURE_SLIDE.generate()
            logger.debug("method complete!")
        case _:
            print("Exception: unsuppotted slide")

print("loops complete!")