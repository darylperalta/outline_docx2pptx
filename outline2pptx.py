from pptx import Presentation

'''
prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('test.pptx')
'''

prs = Presentation('try.pptx')
print('prs', prs)
for slide in prs.slides:
    print('slides: ', len(prs.slides))
    print('slide shapes shape',len(slide.shapes))
    for shape in slide.shapes:
        print('paragraph', len(shape.text_frame.paragraphs))
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            print('runs ', len(paragraph.runs))
            for run in paragraph.runs:
                # text_runs.append(run.text)
                print(run.text)

print('asdf', prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text)
# print(prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text)
prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text = 'asdfkjasfd'
prs.save('try2.pptx')
