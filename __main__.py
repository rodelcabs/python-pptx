import pythonpptx

contentList = pythonpptx.getSlideTextList("content.txt")
ppt = pythonpptx.createPPT("bg1.jpg", "Kay Ganda Ng Umaga", "sample_ppt.pptx", contentList, True)

print(contentList)

