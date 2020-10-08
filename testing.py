from picsheet import make_sheet_from_pic


path = 'successkidmeme.jpg'
name = path.split('.')[0]

make_sheet_from_pic(path, name+'.xlsx')