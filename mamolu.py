import win32gui
import win32con
import win32api
import win32com.client
import time
import os
import sys

target_dir=r"D:\AllDowns\newbooks"

# def acrobat_extract_text(f_path, f_path_out, f_basename, f_ext):
#     avDoc = win32com.client.Dispatch("AcroExch.AVDoc") # Connect to Adobe Acrobat
#
#     # Open the input file (as a pdf)
#     ret = avDoc.Open(f_path, f_path)
#     assert(ret) # FIXME: Documentation says "-1 if the file was opened successfully, 0 otherwise", but this is a bool in practise?
#
#     pdDoc = avDoc.GetPDDoc()
#
#     dst = os.path.join(f_path_out, ''.join((f_basename, f_ext)))
#
#     # Adobe documentation says "For that reason, you must rely on the documentation to know what functionality is available through the JSObject interface. For details, see the JavaScript for Acrobat API Reference"
#     jsObject = pdDoc.GetJSObject()
#
#     # Here you can save as many other types by using, for instance: "com.adobe.acrobat.xml"
#     jsObject.SaveAs(dst, "com.adobe.acrobat.accesstext")
#
#     pdDoc.Close()
#     avDoc.Close(True) # We want this to close Acrobat, as otherwise Acrobat is going to refuse processing any further files after a certain threshold of open files are reached (for example 50 PDFs)
#     del pdDoc



def get_hd_from_child_hds(father_hd,some_idx,expect_name):
    child_hds=[]
    win32gui.EnumChildWindows(father_hd,lambda hwnd, param: param.append(hwnd),child_hds)

    names=[win32gui.GetWindowText(each) for each in child_hds]
    hds=[hex(each) for each in child_hds]
    print("ChildName List:",names)
    print("Child Hds List:",hds)

    name=names[some_idx]
    hd=hds[some_idx]

    print("The {} Child.".format(some_idx))
    print("The Name:{}".format(name))
    print("The HD:{}".format(hd))

    if name==expect_name:
        return child_hds[some_idx]
    else:
        print("窗口不对！")
        return None

def tt():
    # import os
    import winerror
    from win32com.client.dynamic import Dispatch, ERRORS_BAD_CONTEXT

    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)

    my_dir = r"D:\AllDowns"
    my_pdf = "gpgp.pdf"

    os.chdir(my_dir)
    src = os.path.abspath(my_pdf)

    try:
        AvDoc = Dispatch("AcroExch.AVDoc")

        if AvDoc.Open(src, ""):
            pdDoc = AvDoc.GetPDDoc()
            jsObject = pdDoc.GetJSObject()
            jsObject.SaveAs(os.path.join(my_dir, 'output_example.jpeg'), "com.adobe.acrobat.jpeg")

    except Exception as e:
        print(str(e))

    finally:
        AvDoc.Close(True)

        jsObject = None
        pdDoc = None
        AvDoc = None

def renamer(type_num,pdf_name):
    old_name=pdf_name
    new_name=f"typetype{type_num}{old_name}"
    os.rename(f"{target_dir}{os.sep}{old_name}",f"{target_dir}{os.sep}{new_name}")
    print("rename done.")

def check_type(pdf_path,pdf_name,pdf_dir=target_dir):
    # pdf_path="D:\\AllDowns\\gpgp.pdf"

    # pdf_name="gpgp.pdf"
    #
    # os.chdir(pdf_dir)
    # pdf_name=os.path.abspath(pdf_name)
    #
    # avDoc=win32com.client.DispatchEx("AcroExch.AVDoc")
    # ret = avDoc.Open(f"\"{pdf_path}\"",f"\"{pdf_path}\"")
    #
    # time.sleep(5)

    # 用com去控制老是有问题，只能改成默认设置妈的！

    # adobe_comm=f"D:\\Acrobat 11.0\\Acrobat\\AcroRd32.exe /n \"{pdf_path}\""
    #
    # print(adobe_comm)

    # os.system(adobe_comm)

    # 用com去控制老是有问题，唉...

    # nmd，每次运行之前都要查看一下用的是哪款pdf阅读器就离谱...
    # 2020年9月26日01:57:50，重启试了下发现还是gg。算了算了..

    os.startfile(pdf_path)

    time.sleep(2)

    adobe_str=f"{pdf_name} - Adobe Acrobat Pro"

    root_hd=None

    adobe_hd=win32gui.FindWindowEx(root_hd,0,0,adobe_str)

    avui_CommandWidget_hd=get_hd_from_child_hds(adobe_hd,0,expect_name="AVUICommandWidget")

    # 固定是第二个...
    state_hd=get_hd_from_child_hds(avui_CommandWidget_hd,1,expect_name="")

    # https://blog.csdn.net/qq_41928442/article/details/88937337

    # 获取识别结果中输入框文本
    length = win32gui.SendMessage(state_hd, win32con.WM_GETTEXTLENGTH)+1
    buf = win32gui.PyMakeBuffer(length)
    #发送获取文本请求
    win32api.SendMessage(state_hd, win32con.WM_GETTEXT, length, buf)
    #下面应该是将内存读取文本
    address, length = win32gui.PyGetBufferAddressAndLen(buf[:-1])
    text = win32gui.PyGetString(address, length)

    flag=0
    print(f"Text:{text}")
    if text=="1":
        print("Type 1")
        flag=1
    elif text=="A":
        print("Type 2")
        flag=2

    win32gui.PostMessage(adobe_hd, win32con.WM_CLOSE,0,0)
    print("one done.")

    # avDoc.Close(True)
    #
    # del avDoc

    time.sleep(2)

    return flag

def main():
    for each in os.listdir(target_dir):
        if each.endswith(".pdf"):
            pdf_path=os.path.join(target_dir,each)
            pdf_path2=f"{target_dir}{os.sep}{each}"
            print(f"Path:{pdf_path}")
            type_num=check_type(pdf_path,each)
            renamer(type_num,each)
    print("all done.")

if __name__ == '__main__':
    main()
    # tt()
