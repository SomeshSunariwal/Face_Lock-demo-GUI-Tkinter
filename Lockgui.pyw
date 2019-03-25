import tkinter
import tkinter.messagebox as tmsg
from PIL import ImageTk, Image
import cv2
import os
import xlrd
from xlutils.copy import copy
import time
import numpy as np

# Number of samples
number_sample = 100


# Setting up all the definition
def folder():
    if not (os.path.exists("dataset") or os.path.exists("dataset_temp") or os.path.exists("trainer")):
        os.mkdir("dataset")
        os.mkdir("dataset_temp")
        os.mkdir("trainer")
        os.mkdir("dataset_temp/train")
        os.mkdir("dataset_temp/rec")
        time.sleep(2)
        tmsg.showinfo("Success-full", "Folder Successfully created !")
    else:
        pass


def help_msg():
    tmsg.showinfo("Help", """
    Help menu for Face Lock System
    
    #### ADD FACE ####
    
    1.  First Click on the Add New Face
    2.  Fill the required detail which is shown 
        in the dialog box
    3.  Then Hit the Register Button
        and wait for register
        
    #### Recognize ####
    
    1. Click on Recognize Face
    2. Wait for the Response
    
    ## Copyright By Somesh Sunariwal
    """)


def execute():
    val = tmsg.askyesno("Answer", " Are You Sure ?")
    if val:
        frame_2 = tkinter.Frame(add_win)
        global number_sample
        count, i, flag = 0, 0, 0
        cap = cv2.VideoCapture(0)
        value_1 = book2.nrows
        print("Exe  ", value_1)
        for check_ID in range(1, value_1):
            if int(user_id.get()) == int(book2.cell_value(check_ID, 1)) or user_nm.get().capitalize() == str(book2.cell_value(check_ID, 0)):
                flag = 1
                break
            else:
                flag = 0
        if flag == 1:
            tmsg.showerror("Error", "Error! Already Exist Username UserID")

        else:
            while count <= number_sample:
                ret, frame = cap.read()
                frame = cv2.resize(frame, (300, 300))
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = detector.detectMultiScale(gray, 1.3, 5)
                for (x, y, w, h) in faces:
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
                    count += 1
                    cv2.imwrite("dataset/" + user_nm.get() + "." + str(user_id.get()) + '.' + str(count) + ".jpg",
                                gray[y:y + h, x:x + w])
                cv2.imwrite(f"dataset_temp/train/{count}.jpg", frame)
                if cv2.waitKey(25) & 0xff == ord("q"):
                    break
                time.sleep(0.1)
                image = Image.open(f"dataset_temp/train/{count}.jpg")
                image = ImageTk.PhotoImage(image)
                tkinter.Label(frame_2, image=image, pady=10, padx=20).grid(row=0, column=0)
                frame_2.pack(side="bottom")
                os.remove(f"dataset_temp/train/{count}.jpg")
                add_win.update()
            book4.write(value_1, 0, user_nm.get().capitalize())
            book4.write(value_1, 1, int(user_id.get()))
            book3.save("ID.xls")
            tmsg.showinfo("Done", "Your Face is Added to the System")
            var.set("Saving....")
            s_bar.update()
            time.sleep(2)
            var.set("Ready...")
            s_bar.update()
            time.sleep(2)
            add_win.destroy()
            cap.release()
            cv2.destroyAllWindows()
            train("dataset/")
            add_win.mainloop()
    else:
        pass


def train(path):
    var.set("Training Data... Please wait....")
    s_bar.update()
    image_paths = [os.path.join(path, f) for f in os.listdir(path)]
    face_samples = []
    ids = []
    for imagePath in image_paths:
        pil_img = Image.open(imagePath).convert('L')
        img_numpy = np.array(pil_img, 'uint8')
        id = int(os.path.split(imagePath)[-1].split(".")[1])
        faces = detector.detectMultiScale(img_numpy)
        for (x, y, w, h) in faces:
            face_samples.append(img_numpy[y:y + h, x:x + w])
            ids.append(id)
    print(ids)
    recognizer.train(face_samples, np.array(ids))
    recognizer.save('trainer/trainer.yml')
    face_samples.clear()
    ids.clear()
    print("Save Face")
    time.sleep(2)
    var.set("Saving Train Data")
    s_bar.update()
    time.sleep(2)
    var.set("Ready..")
    s_bar.update()
    tmsg.showinfo("Done", "Success Full ! Ready to Rock and Roll")


def win_new(width, height, title, num):
    global add_win
    add_win = tkinter.Toplevel()
    add_win.title(title)
    win_width_ad = width
    win_height_ad = height
    scr_width_ad = add_win.winfo_screenwidth()
    scr_height_ad = add_win.winfo_screenheight()
    add_win.geometry(f"{win_width_ad}x{win_height_ad}+{int(((scr_width_ad / 2) - (win_width_ad / 2)))}+{int(((scr_height_ad / 2) - (win_height_ad / 2)))}")
    add_win.resizable(width=False, height=False)
    if num == 1:
        frame_1 = tkinter.Frame(add_win)
        tkinter.Label(frame_1, text="User Name", font="Times 10", pady=7).grid(row=1, column=1)
        tkinter.Label(frame_1, text="User ID", font="Times 10", pady=7).grid(row=2, column=1)
        tkinter.Entry(frame_1, textvariable=user_nm).grid(row=1, column=2)
        tkinter.Entry(frame_1, textvariable=user_id).grid(row=2, column=2)
        tkinter.Label(frame_1, text="                 ").grid(row=4, column=2)
        tkinter.Button(frame_1, text="Register Now !", command=execute, width=14).grid(row=3, columnspan=3)
        frame_1.pack(side="top")


def add_fc():
    win_new(400, 400, "Add Face", 1)


def user():
    add_win2 = tkinter.Toplevel()
    add_win2.title("Detail of users")
    add_win2.geometry("400x380")
    scrollbar = tkinter.Scrollbar(add_win2)
    scrollbar.pack(fill="y", side="right")
    list_1 = tkinter.Listbox(add_win2)
    list_1.pack(fill="both")
    for i in range(1, book2.nrows):
        data = book2.cell_value(i, 0)
        list_1.insert("end", f"{i}. {data}")
    scrollbar.config(command=list_1.yview)
    add_win2.mainloop()


def rec_fc():
    global flag_1, value, value_2
    i, flag_1 = 0, 0
    recognizer.read("trainer/trainer.yml")
    win_new(400, 400, "Recognize Face", 2)
    cam = cv2.VideoCapture(0)
    book_1 = xlrd.open_workbook("ID.xls")
    book_2 = book_1.sheet_by_index(0)
    value_2 = book_2.nrows
    print("rec : ", value_2)
    var.set("Recognize Working...")
    s_bar.update()
    while i < 50:
        ret, im = cam.read()
        gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        faces = detector.detectMultiScale(gray, 1.2, 5)
        for (x, y, w, h) in faces:
            cv2.rectangle(im, (x - 20, y - 20), (x + w + 20, y + h + 20), (0, 255, 0), 4)
            Id = recognizer.predict(gray[y:y + h, x:x + w])
            holder = Id[0]
            pride = Id[1]

            for check_id in range(1, value_2):
                if holder == int(book2.cell_value(check_id, 1)) and pride >= 65:
                    Id = str(book2.cell_value(check_id, 0))
                    print(Id)
                    break
                else:
                    Id = 'Unknown'
            # cv2.rectangle(im, (x - 22, y - 90), (x + w + 22, y - 22), (0, 255, 0), 1)
            cv2.putText(im, str(Id), (x, y - 40), font, 2, (255, 255, 255), 3)
            label = tkinter.Label(add_win, text=str(Id), width=27, font="Times 20")
            label.grid(row=0, column=0)
            label.update()
            if str(Id) == "Unknown":
                flag_1 = 0
                value = "Unknown"
            else:
                flag_1 = 1
                value = Id
        im = cv2.resize(im, (380, 380))
        cv2.imwrite(f"dataset_temp/rec/{i}.png", im)
        photo_old = Image.open(f"dataset_temp/rec/{i}.png")
        photo_old = ImageTk.PhotoImage(photo_old)
        tkinter.Label(add_win, image=photo_old).grid(row=1, column=0)
        os.remove(f"dataset_temp/rec/{i}.png")
        add_win.update()
        i += 1
        if cv2.waitKey(25) & 0xFF == ord('q'):
            break
    time.sleep(2)
    var.set("Ready..")
    s_bar.update()
    cv2.destroyAllWindows()
    cam.release()
    if flag_1 == 1:
        tmsg.showinfo("Found Match", f"Found {value} ! Have Fun")
    else:
        tmsg.showerror("Error", "Error! Found No Identity Try Again!")
    add_win.destroy()


if __name__ == '__main__':
    # Setting up the XL sheet
    book = xlrd.open_workbook("ID.xls")
    book2 = book.sheet_by_index(0)
    book3 = copy(book)
    book4 = book3.get_sheet("sheet1")
    book4.write(0, 1, "Customer_ID")
    book4.write(0, 0, "customer_name")
    book3.save("ID.xls")
    # Recognizer
    detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    font = cv2.FONT_HERSHEY_SIMPLEX
    # window Create
    window = tkinter.Tk()
    window.title("Face Lock System")
    # Setup Window
    win_width = 400
    win_height = 500
    scr_width = window.winfo_screenwidth()
    scr_height = window.winfo_screenheight()
    window.geometry(f"{win_width}x{win_height}+{int(((scr_width/2)-(win_width/2)))}+{int(((scr_height/2)-(win_height/2)))}")
    window.resizable(width=False, height=False)
    # Setting up Variable
    var = tkinter.StringVar()
    var.set("Ready....")
    user_nm = tkinter.StringVar()
    user_id = tkinter.StringVar()
    # Setting up the menu system
    main_menu = tkinter.Menu(window)
    # File menu setup
    File_menu = tkinter.Menu(main_menu, tearoff=0)
    main_menu.add_cascade(label="File", menu=File_menu)
    File_menu.add_command(label="Registered Users", command=user)
    File_menu.add_command(label="Folder", command=folder)
    File_menu.add_separator()
    File_menu.add_command(label="Quit", command=quit)
    # Help menu setup
    help_menu = tkinter.Menu(main_menu, tearoff=0)
    main_menu.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Help", command=help_msg)
    window.config(menu=main_menu)
    # Setting up the image in window gui
    photo = Image.open("1.jpg")
    photo = ImageTk.PhotoImage(photo)
    tkinter.Label(window, image=photo).pack(anchor="n", side="top")
    tkinter.Label(window, text="Copyright @ Somesh Sunariwal 2019", relief="sunken").pack(fill="x", side="bottom")
    s_bar = tkinter.Label(window, textvariable=var, relief="sunken", anchor="w")
    s_bar.pack(fill="x", side="bottom")
    tkinter.Button(window, text="Add Face", command=add_fc, pady=12, width=17, height=12, font="IMPACT 18").pack(side="left")
    tkinter.Button(window, text="Recognize Face", command=rec_fc, pady=12, width=17, height=12, font="IMPACT 18").pack(side="left")

    window.mainloop()

