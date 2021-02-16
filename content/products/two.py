import codecs
import openpyxl

book = openpyxl.open("troubleoff.xlsx", read_only=True)

sheet = book.active



for row in range(2, sheet.max_row+1):
    idd = sheet[row] [0].value
    title = sheet[row] [1].value
    description = sheet[row] [2].value
    price = sheet[row] [3].value
    order = sheet[row] [4].value
    default_original_image = sheet[row] [5].value
    default_thumbnail_image = sheet[row] [6].value
    category = sheet[row] [7].value
    featured = sheet[row] [8].value
    layout = sheet[row] [9].value

    file_name = 'file{}.md'.format(row-1)
    with codecs.open(file_name,'w', 'utf-8') as f:
        #    f.write("id:" + repr(str(idd)) + "\n" +
        #    "title:" + title + "\n" +
        #    "description:" + description + "\n" +
        #    "price:" + repr(str(price)) + "\n" +
        #    "order:" +  order + "\n" + 
        #    "default_thumbnail_image:" +  default_thumbnail_image + "\n" + 
        #    "default_original_image:" +  default_original_image + "\n" + 
        #    "category:" + category + "\n" + 
        #    "layout:" + layout)


        f.write("---\nid: '{}'\ntitle: {}\ndescription: {}\nprice: '{}'\norder: {}\ndefault_thumbnail_image: {}\ndefault_original_image: {}\ncategory: {}\nfeatured: {}\nlayout: {}\n---".format(idd, title, description, price, order, default_thumbnail_image, default_original_image, category, featured, layout))