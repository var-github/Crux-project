import pypdf
from math import cos, atan, pi
from spire.pdf import *
from docx2pdf import convert as doc_convert
from pptxtopdf import convert as ppt_convert


def take_pages_input(prompt, no_of_pages):
    inp = input(prompt).strip().split(",")
    i = 0
    while i < len(inp):
        if "-" not in inp[i]:
            try:
                inp[i] = int(inp[i])
            except:
                print(f"'{inp[i]}' is not a valid page number!!\n\n")
                break
            if inp[i] < 1 or inp[i] > no_of_pages:
                print(f"Page '{inp[i]}' does not exist in the file!!\n\n")
                break
        else:
            if inp[i].count("-") != 1:
                print(f"'{inp[i]}' is an invalid page range!!\n\n")
                break
            try:
                inp[i] = [int(inp[i].split("-")[0]), int(inp[i].split("-")[1])]
            except:
                print(f"'{inp[i]}' is an invalid page range!!\n\n")
                break
            if inp[i][0] < 1 or inp[i][0] > no_of_pages or inp[i][1] < inp[i][0]:
                print(f"'{inp[i][0]} - {inp[i][1]}' is not a valid page range!!\n\n")
                break
            if inp[i][1] > no_of_pages:
                inp[i][1] = no_of_pages
            n = inp[i][1] - inp[i][0]
            inp[i:i + 1] = list(range(inp[i][0], inp[i][1] + 1))
            i += n
        i += 1
    else:
        return inp
    return None


def take_int_input(prompt, error_message, lower_limit=1, upper_limit=0):
    n = input(prompt).strip()
    try:
        n = int(n)
        if n < lower_limit or (upper_limit and n > upper_limit):
            1 / 0
    except:
        print(error_message.format(n))
        return None
    return n


def take_path_input(raw=False):
    p = input("\nEnter path to pdf: ").strip().strip('"')
    try:
        pdf = pypdf.PdfReader(p)
    except:
        print("File does not exist or file is not pdf!!\n\n")
        return None, None
    if not raw and pdf.is_encrypted:
        print("File is encrypted, first remove password before using other tools!!\n\n")
        return None, None
    return p, pdf


def merge_pdf():
    f = pypdf.PdfWriter()
    n = take_int_input("\nEnter how many pdf's you want to merge: ", "Cant merge '{}' pdf's!!\n\n", lower_limit=2)
    if not n:
        return
    for i in range(n):
        name = input("Enter bookmark to add at the beginning of this pdf (leave empty for no bookmark): ")
        p = input(f"Enter path to pdf {i + 1}: ").strip().strip('"')
        try:
            if name:
                f.append(p, outline_item=name)
            else:
                f.append(p)
        except:
            print("File does not exist!!\n\n")
            f.close()
            return
    p = input("\nEnter path/name to give merged pdf: ").strip(".pdf") + ".pdf"
    f.write(p)
    f.close()
    print("PDF merged successfully!!\n\n")


def split_pdf():
    p, pdf = take_path_input()
    if not p:
        return
    no_of_pages = len(pdf.pages)
    n = take_int_input("Enter number of pdf's to split into: ", "Cant do '{}' splits on a pdf!!\n\n")
    if not n:
        return
    for i in range(n):
        print("Split", i + 1)
        pages = input("Enter page range (e.g. 10-20, both 10 and 20 are included): ")
        if pages.count("-") != 1:
            print("Please enter a valid page range!!\n\n")
            return
        try:
            pages = list(map(int, pages.split("-")))
        except:
            print("Please enter only integer ranges!!\n\n")
            return
        if no_of_pages < pages[0] or pages[0] < 1 or no_of_pages < pages[1] or pages[1] < pages[0]:
            print("Please enter a valid page range!!\n\n")
            return
        split = pypdf.PdfWriter()
        split.append(pdf, pages=(pages[0]-1, pages[1]-1))
        split.write(p[:-4] + "_split" + str(i + 1) + ".pdf")
        print("Split", i + 1, "successful, stored in original directory")
    print("\n")


def reorganise():
    p, pdf = take_path_input()
    if not p:
        return
    no_of_pages = len(pdf.pages)
    print("\nInstructions\nEnter page number/ranges in required order separated by comma")
    print("Any page numbers missing will be deleted")
    print("Example input: 1, 3-6, 4-8, 9")
    print("The above input will generate a pdf with pages in this order 1,3,4,5,6,4,5,6,7,8,9\n")
    pages = take_pages_input("Enter page numbers/ranges: ", no_of_pages)
    if not pages:
        return
    reorder = pypdf.PdfWriter()
    for i in pages:
        reorder.add_page(pdf.pages[i - 1])
    reorder.write(p[:-4] + "_reordered.pdf")
    print("PDF reordered successfully!!\n\n")


def rotate_pages():
    p, pdf = take_path_input()
    if not p:
        return
    no_of_pages = len(pdf.pages)
    angle = input("Enter angle in degrees to rotate by: ").strip()
    try:
        angle = int(angle)
    except:
        print("Angle must be an integer!!\n\n")
        return
    pages = take_pages_input(f"Enter page numbers/ranges that you want to rotate by {angle} degrees: ", no_of_pages)
    if not pages:
        return
    output = pypdf.PdfWriter()
    for i in range(no_of_pages):
        output.add_page(pdf.pages[i])
        if i + 1 in pages:
            x0 = (output.pages[i].mediabox.right - output.pages[i].mediabox.left) / 2
            y0 = (output.pages[i].mediabox.top - output.pages[i].mediabox.bottom) / 2
            s0 = min(x0, y0) / cos(atan(max(x0, y0) / min(x0, y0)) - abs(((angle - 90) * pi / 180) % pi - pi / 2)) / (
                        x0 ** 2 + y0 ** 2) ** 0.5
            output.pages[i].add_transformation(
                pypdf.Transformation().translate(tx=-x0, ty=-y0).scale(sx=s0, sy=s0).rotate(angle).translate(tx=x0,
                                                                                                             ty=y0))
    output.write(p[:-4] + "_rotated.pdf")
    print("Pages successfully rotated!!\n\n")


def encrypt_or_decrypt():
    p, pdf = take_path_input(raw=True)
    if not p:
        return
    if pdf.is_encrypted:
        pdf.decrypt(input("Enter password (to remove it): "))
        try:
            pdf = pypdf.PdfWriter(clone_from=pdf)
        except:
            print("Wrong password!!\n\n")
            return
        print("Password removed successfully\n\n")
    else:
        pdf = pypdf.PdfWriter(clone_from=pdf)
        pdf.encrypt(input("Enter password to add: "), algorithm="AES-256")
        print("Password added successfully\n\n")
    pdf.write(p)


def add_bookmark():
    p, pdf = take_path_input()
    if not p:
        return
    pdf = pypdf.PdfWriter(clone_from=p)
    no_of_pages = len(pdf.pages)
    page = take_int_input("Enter page to add bookmark to: ", "Page '{}' does not exist in the file!!\n\n",
                          upper_limit=len(pdf.pages))
    if not page:
        return
    parent = pdf.add_outline_item(title=input("Enter bookmark title: "), page_number=page - 1)
    if input("Do you want to add sub bookmarks?(y/n): ").lower().strip()[0] == 'y':
        n = take_int_input("Enter how many sub bookmarks you want to add: ", "Cant add '{}' sub bookmarks!!\n\n")
        if not n:
            return
        for i in range(n):
            pdf.add_outline_item(title=input("Enter bookmark title: "), page_number=page - 1, parent=parent)
    p = p.rstrip("_bookmark.pdf").rstrip(".pdf")
    pdf.write(p + "_bookmark.pdf")
    print("Bookmark successfully added!!\n\n")


def edit_bookmark():
    p = input("\nEnter path to pdf: ").strip().strip('"')
    try:
        pdf = PdfDocument()
        pdf.LoadFromFile(p)
    except:
        print("File does not exist!!\n\n")
        return
    bookmarks = pdf.Bookmarks
    if bookmarks.Count > 0:
        print("Pdf bookmarks:")
        for i in range(bookmarks.Count):
            bookmark = bookmarks.get_Item(i)
            print(f"{i + 1}. {bookmark.Title}")
            child = PdfBookmarkCollection(bookmark)
            if child.Count > 0:
                for j in range(child.Count):
                    childBookmark = child.get_Item(j)
                    print(f"  {i + 1}.{j + 1}. {childBookmark.Title}")
    print("Remember - To edit sub-bookmark enter its parent bookmarks's name")
    no = take_int_input("\nEnter bookmark number to edit: ", "Bookmark number '{}' does not exist in the file!!\n\n",
                        upper_limit=bookmarks.Count)
    if not no:
        return
    bookmark = bookmarks.get_Item(no - 1)
    child = PdfBookmarkCollection(bookmark)
    child.Clear()
    page = take_int_input("Enter page number to add bookmark to: ", "Page '{}' does not exist in the file!!\n\n",
                          upper_limit=pdf.Pages.Count)
    if not page:
        return
    bookmark.Title = input("Enter bookmark title: ")
    bookmark.Action = PdfGoToAction(PdfDestination(pdf.Pages[page - 1]))
    if input("Do you want to add sub bookmarks?(y/n): ").lower().strip()[0] == 'y':
        n = take_int_input("Enter how many sub bookmarks you want to add: ", "Cant add '{}' sub bookmarks!!\n\n")
        if not n:
            return
        bookmark_collection = PdfBookmarkCollection(bookmark)
        for i in range(n):
            childBookmark = bookmark_collection.Add(input(f"Enter tite for sub bookmark {i + 1}: "))
            childBookmark.Action = PdfGoToAction(PdfDestination(pdf.Pages[page - 1]))
    p = p.rstrip("_bookmark.pdf").rstrip(".pdf")
    pdf.SaveToFile(p + "_bookmark.pdf")
    print("Bookmark successfully added!!\n\n")


def delete_bookmark():
    p = input("\nEnter path to pdf: ").strip().strip('"')
    try:
        pdf = PdfDocument()
        pdf.LoadFromFile(p)
    except:
        print("File does not exist!!\n\n")
        return
    bookmarks = pdf.Bookmarks
    if bookmarks.Count > 0:
        print("Pdf bookmarks:")
        for i in range(bookmarks.Count):
            bookmark = bookmarks.get_Item(i)
            print(f"{i + 1}. {bookmark.Title}")
            child = PdfBookmarkCollection(bookmark)
            if child.Count > 0:
                for j in range(child.Count):
                    childBookmark = child.get_Item(j)
                    print(f"  {i + 1}.{j + 1}. {childBookmark.Title}")
    no = take_int_input("Enter bookmark number to delete: ", "Bookmark number '{}' does not exist in the file!!\n\n",
                        upper_limit=bookmarks.Count)
    if not no:
        return
    pdf.Bookmarks.RemoveAt(no - 1)
    p = p.rstrip("_bookmark.pdf").rstrip(".pdf")
    pdf.SaveToFile(p + "_bookmark.pdf")
    print("Bookmark successfully removed!!\n\n")


def scale_pages():
    p, pdf = take_path_input()
    if not p:
        return
    pdf = pypdf.PdfWriter(clone_from=p)
    no_of_pages = len(pdf.pages)
    size = input("Scale pages to which size (A0, A1, A2, A3, A4, A5, A6, A7, A8, C4): ").strip().upper()
    if size not in ["A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "C4"]:
        print("Invalid size!!\n\n")
        return
    pages = take_pages_input(f"Enter page numbers/ranges that you want to scale: ", no_of_pages)
    if not pages:
        return
    for i in pages:
        pdf.pages[i - 1].scale_to(eval(f"pypdf.PaperSize.{size}.height"), eval(f"pypdf.PaperSize.{size}.width"))
    pdf.write(p.rstrip("_scaled.pdf").rstrip(".pdf") + "_scaled.pdf")
    print("Page successfully scaled!!\n\n")


while True:
    print("PDF handling".center(50))
    print("1. Merge pdf\n2. Split\n3. Reorganise\n4. Compress PDF\n5. Rotate pages\n6. Add/Remove password")
    print("7. Add bookmark\n8. Edit bookmark\n9. Delete bookmark\n10. Scale pages\n11. Convert to pdf\n12. Extract text")
    print("13. Exit")
    ch = input("Enter choice: ")
    try:
        ch = int(ch)
    except:
        print("Invalid option\n\n")
        continue
    if ch == 1:
        merge_pdf()
    elif ch == 2:
        split_pdf()
    elif ch == 3:
        reorganise()
    elif ch == 4:
        p, pdf = take_path_input()
        if not p:
            continue
        pdf = PdfCompressor(p)
        pdf.CompressToFile(p[:-4] + "_compressed.pdf")
        print("File successfully compressed, compressed file is in original directory\n\n")
    elif ch == 5:
        rotate_pages()
    elif ch == 6:
        encrypt_or_decrypt()
    elif ch == 7:
        add_bookmark()
    elif ch == 8:
        edit_bookmark()
    elif ch == 9:
        delete_bookmark()
    elif ch == 10:
        scale_pages()
    elif ch == 11:
        p = input("Enter path to file that is to be converted to pdf: ").strip().strip('"')
        try:
            doc_convert(p, p.rstrip(".docx") + ".pdf")
        except:
            try:
                ppt_convert(p, p[:p.rfind("\\")])
            except:
                print("Unable to convert!!\n\n")
                continue
        print("File successfully converted!!\n\n")
    elif ch == 12:
        p, pdf = take_path_input()
        if not p:
            continue
        pdf = pypdf.PdfReader(p)
        f = open(p[:-4]+".txt", "wb")
        for page in pdf.pages:
            f.write(str(page.extract_text(extraction_mode="layout")).encode())
        f.close()
        print("Text successfully extracted!!\n\n")
    elif ch == 13:
        print("Exiting...")
        break
    else:
        print("Invalid option\n\n")
