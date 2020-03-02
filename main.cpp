#include <windows.h>
#include <cstdlib>
#include <iostream>
#include <conio.h>
#include <Commdlg.h>
#include "libxl.h"
//#include "resource.h"
using namespace libxl;

#define ID_TIMER    1
#define SWITCH      0
#define FILE1       1
#define FILE2       2
#define ID_LIST_1     3
#define ID_LIST_2     4
struct
{
     int     iStyle ;
     TCHAR * szText ;
}
button[10] =
{
     BS_PUSHBUTTON, TEXT ("��ʼ") ,
	 BS_PUSHBUTTON, TEXT ("��������ļ�"),
	 BS_PUSHBUTTON, TEXT ("���ƥ���ļ�")   
} ;
// ����ֵ: �ɹ� 1, ʧ�� 0
// ͨ�� path ���ػ�ȡ��·��
static char LIST_1_PATH[200],LIST_2_PATH[20][200];//������ļ�·�� 
static char list1[200],list2[20][200]; //�б���ʾ���� 
static int  cnt =0; //�����Ҷ��ٸ��ļ� 
static bool SWITCH_ALL = false; 
int FileDialog(char *path)  //ѡ���ļ��Ի��� 
{
	OPENFILENAME ofn;
	ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn); // �ṹ��С
    ofn.lpstrFile = path; // ·��
    ofn.nMaxFile = MAX_PATH; // ·����С
    ofn.lpstrFilter = "All\0*.*\0Text\0*.TXT\0"; // �ļ�����
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
	return GetOpenFileName(&ofn);
}
const int IND= 2; //Ҫ��ȡ��sheet���� 
static TCHAR   szFormat[] = TEXT ("%-16s%04X-%04X    %04X-%04X"); 
bool searchExcel(TCHAR  * str){
	 //MessageBox(NULL, str,"Error!",MB_ICONEXCLAMATION|MB_OK);
	 bool flag = false;
	 Book* book = xlCreateBook();
	 char NO[300]={"����: "},SUCCESS[30]={" ���ҳɹ�"},FAIL[30]={" ����ʧ��"}; 	 
     try
	 {
	   // ��������
	   	 if(book){
		 	 for(int i=0;i<cnt;i++){
		 	 	if(book->load(LIST_2_PATH[i])){
					Sheet* sheet = book->getSheet(0);
		            if(sheet)  //first - last -1
		            {   	
		                 int first= sheet->firstRow(),last= sheet->lastRow();
		                 for(int j=first;j<last;j++){
		                 	const TCHAR * tmp=sheet->readStr(j,IND);
		                 	if(tmp != NULL && strcmp((TCHAR *)tmp,str)==0) {
		                 		//string s= "���룺" +  (string)tmp  + "����";
		                 		strcat(NO,tmp);strcat(NO,SUCCESS);
								strcat(NO,"\n���ļ�: "); strcat(NO,list2[i]); //NO[strlen(NO)] =(char)j
								char st[100];ltoa(j+1,st,10);
								strcat(NO," �ĵ�_"); strcat(NO,st);    strcat(NO, "_��" );   
								MessageBox(NULL,(TCHAR *)NO,"�ɹ���",MB_ICONEXCLAMATION|MB_OK);
		                 		book->release();
								return true;
		                 	}
		                 }
		            }
				 }
		 	 } 
		}
		book->release();
		strcat(NO,str);strcat(NO,FAIL);
		MessageBox(NULL,(TCHAR *)NO,"ʧ�ܣ�",MB_ICONEXCLAMATION|MB_OK);
	 }catch(... )
	 {
	  	book->release();
	 }
	return false;
} 
static int ExcelLen=0; //���λ�������ж�excel�Ƿ�仯 
static char CURRENTNO[100]; // �ײ�߿���ʾ�ı� 
void checkLastOne(HWND hwnd){
	Book* book = xlCreateBook(); 
	 try
	 {
		 if(book){
			if(book->load(LIST_1_PATH)){
				Sheet* sheet = book->getSheet(0);
	            if(sheet)  //first - last -1
	            {   	
	                 int last= sheet->lastRow();
	                 const TCHAR * lstr=sheet->readStr(last-1,IND);
	                 if(lstr!=NULL){
	                 	strcpy(CURRENTNO,lstr);
	                 	InvalidateRect (hwnd, NULL, TRUE) ;
	                 	searchExcel((TCHAR *)lstr);
	                 }
	            }
			}
		}
		book->release();
     }catch(... )
	 {
	  	book->release();
	 }
}
void checkExcel(HWND hwnd){
	Book* book = xlCreateBook(); 
	 try
	 {
		if(book){
			if(book->load(LIST_1_PATH)){
				Sheet* sheet = book->getSheet(0);
	            if(sheet)  //first - last -1
	            {   	
	                 int last= sheet->lastRow();
	                 if(ExcelLen!=last) {
	                 	int tmp=ExcelLen;
	                 	ExcelLen=last;
	                 	if(tmp!=0)
	                 		checkLastOne( hwnd);
	                 }
	            }
			}
		}
		book->release();
     }catch(... )
	 {
	  	book->release();
	 }

}
void drawLines(HDC hdc,int x,int y,int h,int w)
{
	y+=100;
	MoveToEx(hdc,x,y,NULL);
    LineTo(hdc,x,y+h);
    MoveToEx(hdc,x,y,NULL);
    LineTo(hdc,x+w,y);
    MoveToEx(hdc,x+w,y,NULL);
    LineTo(hdc,x+w,y+h);
    MoveToEx(hdc,x,y+h,NULL);
    LineTo(hdc,x+w,y+h);
}
void FillListBox (HWND hwndList,char str[]) 
{
	SendMessage (hwndList, LB_ADDSTRING, 0, (LPARAM) str) ;
}
void ClearListBox (HWND hwndList) 
{
	SendMessage (hwndList, LB_RESETCONTENT, 0, 0 );
}
#define NUM (sizeof button / sizeof button[0])
/* This is where all the input to the window goes to */
LRESULT CALLBACK WndProc(HWND hwnd, UINT Message, WPARAM wParam, LPARAM lParam) {
	static int   cxChar, cyChar ;
	static HICON hIcon ;
    static int   cxIcon, cyIcon, cxClient, cyClient ;
    HINSTANCE    hInstance ;
	static HWND  hwndButton[NUM],hwndList_2,hwndList_1, hwnd2;
	static RECT  rect ;
	HBRUSH       hBrush ;
	static TCHAR runing[]  = TEXT ("������"),stoped[]  = TEXT ("�ѽ���"),
							 szBuffer[50],
							 checkFile[]=TEXT ("checkFile"),
							 searchFile[]=TEXT ("searchFile"), 
							 szFormat[] = TEXT ("%-16s%04X-%04X    %04X-%04X");
	HDC          hdc ;
    PAINTSTRUCT  ps ;
	int          i,len,len2,index,x,y; 
	static int   flag=1; //��ʼ��������
	char         szFile[MAX_PATH] = {0};
	switch(Message) {
		
		case WM_CREATE :
		  SetTimer (hwnd, ID_TIMER, 1000, NULL) ;
          cxChar = LOWORD (GetDialogBaseUnits ()) ;
          cyChar = HIWORD (GetDialogBaseUnits ()) ;
          hwndButton[0] = CreateWindow (TEXT("button"), 
                               button[0].szText,
                               WS_CHILD | WS_VISIBLE | button[0].iStyle,
                               cxChar, cyChar ,
                               75 * cxChar, 10 * cyChar / 4,
                               hwnd, (HMENU)0,
                               ((LPCREATESTRUCT) lParam)->hInstance, NULL) ;
          hwndButton[1] = CreateWindow (TEXT("button"), 
                               button[1].szText,
                               WS_CHILD | WS_VISIBLE | button[1].iStyle,
                               cxChar , cyChar +50,
                               25 * cxChar+90, 8 * cyChar / 4,
                               hwnd, (HMENU) 1,
                               ((LPCREATESTRUCT) lParam)->hInstance, NULL) ;
          hwndButton[2] = CreateWindow (TEXT("button"), 
                               button[2].szText,
                               WS_CHILD | WS_VISIBLE | button[2].iStyle,
                               cxChar+400-90, cyChar+50 ,
                               25 * cxChar+90, 8 * cyChar / 4,
                               hwnd, (HMENU) 2,
                               ((LPCREATESTRUCT) lParam)->hInstance, NULL) ;
          hwndList_1 = CreateWindow (TEXT ("listbox"), NULL,
                              WS_CHILD | WS_VISIBLE | LBS_STANDARD,
                              cxChar , cyChar +100,
                               25 * cxChar+90,250,
                              hwnd, (HMENU) ID_LIST_1,
                              (HINSTANCE) GetWindowLong (hwnd, GWL_HINSTANCE),
                              NULL) ;
          hwndList_2 = CreateWindow (TEXT ("listbox"), NULL,
                              WS_CHILD | WS_VISIBLE | LBS_STANDARD,
                              cxChar+400-90, cyChar+100 ,
                              25 * cxChar+90, 250,
                              hwnd, (HMENU) ID_LIST_2,
                              (HINSTANCE) GetWindowLong (hwnd, GWL_HINSTANCE),
                              NULL) ;
          hwnd2 = GetDlgItem(hwnd, 0);
   		  SetWindowLong(hwnd2, GWL_STYLE, GetWindowLong(hwnd2, GWL_STYLE) | BS_OWNERDRAW);
		  break;
        case WM_SIZE :
          break;
		case WM_PAINT :
          hdc = BeginPaint (hwnd, &ps) ;
          drawLines(hdc,cxChar,cyChar+255,55,600);
          TextOut (hdc, cxChar+10, cyChar +50+20 +300, "��ȡ���룺", lstrlen ("��ȡ���룺")) ;
          TextOut (hdc, cxChar+90, cyChar +50+20 +300, CURRENTNO, lstrlen (CURRENTNO)) ;  
          EndPaint (hwnd, &ps) ;
          return 0 ;
	    case WM_COMMAND :
          switch(wParam){
          	case SWITCH:
          	  hdc = GetDC (hwndButton[0]) ;
	          flag=-flag;
	          if(flag==-1){
	          	SetWindowText(hwndButton[0],"������...(�������)");//��ʼ
	          	SWITCH_ALL=true; 
	          }else {
	          	SWITCH_ALL=false;//�رռ�� 
	          	ExcelLen=0; //ֹ֮ͣ���ʼ�� ExcelLen
	          	SetWindowText(hwndButton[0],"��ʼ");
	          }
	          ReleaseDC (hwnd, hdc) ;
	          InvalidateRect (hwnd, NULL, TRUE) ;
          	  break;
          	case FILE1:
          		if(FileDialog(szFile))
				{
					ClearListBox(hwndList_1);
 					len =strlen(szFile),len2=0;//index
 					for(i=0;i<len;i++) {
 						if(szFile[i]=='\\') index=i;
 					}
 					list1[len2++]=' ';
 					for(i=index+1;i<len;i++) {
 						list1[len2++]=szFile[i];
 					}
 					list1[len2++]='\0';
 					len2=0;
 					for(i=0;i<len;i++){
 						if(szFile[i]!='\\'){
 							LIST_1_PATH[len2++]=szFile[i];
 						}else {
 							LIST_1_PATH[len2++]='/';
 							LIST_1_PATH[len2++]='/';
 						}
 					}
 					LIST_1_PATH[len2]='\0';
					FillListBox (hwndList_1,list1) ;
					ExcelLen=0; //�����ļ�֮��֮���ʼ�� ExcelLen
				}
          		break;
          	case FILE2:
          		if(FileDialog(szFile))
				{
					len =strlen(szFile),len2=0;
					for(i=0;i<len;i++) {
 						if(szFile[i]=='\\') index=i;
 					}
 					list2[cnt][len2++]=' ';
 					for(i=index+1;i<len;i++) {
 						list2[cnt][len2++]=szFile[i];
 					}
 					list2[cnt][len2++]='\0';
 					len2=0;
 					for(i=0;i<len;i++){
 						if(szFile[i]!='\\'){
 							LIST_2_PATH[cnt][len2++]=szFile[i];
 						}else {
 							LIST_2_PATH[cnt][len2++]='/';
 							LIST_2_PATH[cnt][len2++]='/';
 						}
 					}
 					LIST_2_PATH[cnt][len2]='\0';
 					FillListBox (hwndList_2,list2[cnt]) ;
					cnt +=1;
				}
          		break;
          }
          break ;
		case WM_SYSCOLORCHANGE:
			InvalidateRect (hwnd, NULL, TRUE) ;
			break;
		/* Upon destruction, tell the main thread to stop */
		case WM_TIMER:
			//MessageBeep(-1);
			if(SWITCH_ALL)//����˿�ʼ �ټ�� 
				checkExcel(hwnd);
			break;
		case WM_DESTROY: {
			PostQuitMessage(0);
			break;
		}
		case WM_CTLCOLORBTN :
        	if ((HWND)lParam == GetDlgItem(hwnd, 0))
	        {
	            HWND hwnd2 = (HWND)lParam;
	            HDC hdc = (HDC)wParam;
				RECT rc;
	            TCHAR text[64];
	 
	            GetWindowText(hwnd2, text, 63);
	            GetClientRect(hwnd2, &rc);
	            SetTextColor(hdc, RGB(0, 0, 0));
	            SetBkMode(hdc, TRANSPARENT);
	            DrawText(hdc, text, strlen(text), &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
	 
	            if(flag==1) //154 255 154
	            	return (LONG)CreateSolidBrush(GetSysColor(COLOR_BTNFACE));
	        	else 
	        		return (LONG)CreateSolidBrush(RGB(154,255,154));
	        }
        break;
		/* All other messages (a lot of them) are processed using default procedures */
		default:
			return DefWindowProc(hwnd, Message, wParam, lParam);
	}
	return 0;
}

/* The 'main' function of Win32 GUI programs: this is where execution starts */
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	WNDCLASSEX wc; /* A properties struct of our window */
	HWND hwnd; /* A 'HANDLE', hence the H, or a pointer to our window */
	MSG Msg; /* A temporary location for all messages */

	/* zero out the struct and set the stuff we want to modify */
	memset(&wc,0,sizeof(wc));
	wc.cbSize		 = sizeof(WNDCLASSEX);
	wc.lpfnWndProc	 = WndProc; /* This is where we will send messages to */
	wc.hInstance	 = hInstance;
	wc.hCursor		 = LoadCursor(NULL, IDC_ARROW);
	
	/* White, COLOR_WINDOW is just a #define for a system color, try Ctrl+Clicking it */
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
	wc.lpszClassName = "WindowClass";
//    wc.hIcon         = LoadIcon (hInstance, MAKEINTRESOURCE (IDI_ICON)) ;
//    wc.hIconSm       = LoadIcon (hInstance, MAKEINTRESOURCE (IDI_ICON)) ;
	wc.hIcon		 = LoadIcon(NULL, IDI_APPLICATION); /* Load a standard icon */
	wc.hIconSm     = LoadIcon(NULL, IDI_APPLICATION); /* use the name "A" to use the project icon */

	if(!RegisterClassEx(&wc)) {
		MessageBox(NULL, "Window Registration Failed!","Error!",MB_ICONEXCLAMATION|MB_OK);
		return 0;
	}

	hwnd = CreateWindowEx(WS_EX_CLIENTEDGE,"WindowClass","EXCEL���ҳ���",WS_VISIBLE|WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, /* x */
		CW_USEDEFAULT, /* y */
		640, /* width */
		480, /* height */
		NULL,NULL,hInstance,NULL);

	if(hwnd == NULL) {
		MessageBox(NULL, "Window Creation Failed!","Error!",MB_ICONEXCLAMATION|MB_OK);
		return 0;
	}
	/*
		This is the heart of our program where all input is processed and 
		sent to WndProc. Note that GetMessage blocks code flow until it receives something, so
		this loop will not produce unreasonably high CPU usage
	*/
	while(GetMessage(&Msg, NULL, 0, 0) > 0) { /* If no error is received... */
		TranslateMessage(&Msg); /* Translate key codes to chars if present */
		DispatchMessage(&Msg); /* Send it to WndProc */
	}
	return Msg.wParam;
}
