import pypandoc, os, sys, shutil


def main():
    if len(sys.argv) > 1:
        manual_docx = sys.argv[1]
        if manual_docx.endswith('.docx') or manual_docx.endswith('.DOCX'):
            print('found word document', manual_docx)
        else:
            print('please give me a word doc to convert')
            sys.exit()
    else:
        print('please give me a word doc to convert')
        sys.exit()

    convert_docx_to_md(manual_docx)


def convert_docx_to_md(docx):
    print('\nconverting', docx, 'to markdown...')
    output = pypandoc.convert_file(docx, 'md')
    output_file = os.path.splitext(docx)[0] + '.md'
    with open(output_file, 'w') as f:
        f.write(output)
    print('\ncopying file to this directory...')
    shutil.copy(output_file, './')
    print('\nDONE!')


if __name__ == '__main__':
    main()
