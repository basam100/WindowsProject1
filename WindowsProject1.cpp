#include "framework.h"
#include "GradeCalculator.h"
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <commdlg.h>
#include <commctrl.h>
#include <windowsx.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define MAX_LOADSTRING 100
#define ID_EDIT_FILENAME 101
#define ID_BUTTON_LOAD 102
#define ID_BUTTON_EXPORT 103
#define ID_LISTVIEW 104
#define ID_STATUSBAR 105
using namespace std;

HINSTANCE hInst;                              
WCHAR szTitle[MAX_LOADSTRING];                 
WCHAR szWindowClass[MAX_LOADSTRING];          
HWND hWndListView;                           
HWND hWndStatus;                              
HWND hWndEditFilename;                        
HWND hWndButtonLoad;                          
HWND hWndButtonExport;                         

struct Student {
    string firstName;
    string lastName;
    double attendance;
    double groupWork;
    double quizzes[4];
    double labs[8];
    double homework[4];
    double midterm;
    double finalExam;
    double avgQuizzes;
    double avgLabs;
    double avgHomework;
    double courseAverage;
    string letterGrade;
};

vector<Student> students;


ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

string calculateLetterGrade(double avg) {
    if (avg > 90) {
        return "A";
    }
    else if (avg > 80) {
        return "B";
    }
    else if (avg > 75) {
        return "C+";
    }
    else if (avg > 70) {
        return "C";
    }
    else if (avg > 60) {
        return "D";
    }
    return "F";
}

double findMinQuiz(double q[4]) {
    double min = q[0];
    for (int i = 1; i < 4; i++) {
        if (q[i] < min) {
            min = q[i];
        }
    }
    return min;
}

double findAverageQuizScore(double q[4]) {
    double minQuiz = findMinQuiz(q);
    return (q[0] + q[1] + q[2] + q[3] - minQuiz) / 3.0;
}

double findAverageHomeworkScore(double hw[4]) {
    double hw1Percent = (hw[0] / 10.0) * 100.0;
    double hw2Percent = (hw[1] / 10.0) * 100.0;
    double hw3Percent = (hw[2] / 20.0) * 100.0;
    double hw4Percent = (hw[3] / 20.0) * 100.0;

    double averagePercent = (hw1Percent + hw2Percent + hw3Percent + hw4Percent) / 4.0;
    return averagePercent / 10.0;
}

double findAverageLabScore(double arr[8]) {
    return (arr[0] + arr[1] + arr[2] + arr[3] + arr[4] + arr[5] + arr[6] + arr[7]) / 8.0;
}

void InitListView(HWND hWnd) {
    hWndListView = CreateWindowEx(
        0,                      
        WC_LISTVIEW,            
        L"",                   
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL,  
        0, 60,                 
        0, 0,                 
        hWnd,                  
        (HMENU)ID_LISTVIEW,    
        hInst,                
        NULL);                 


    ListView_SetExtendedListViewStyle(hWndListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    LVCOLUMN lvc = { 0 };
    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvc.fmt = LVCFMT_LEFT;

    struct {
        const wchar_t* name;
        int width;
    } columns[] = {
        { L"First Name", 100 },
        { L"Last Name", 100 },
        { L"Quiz Avg", 70 },
        { L"Lab Avg", 70 },
        { L"HW Avg", 70 },
        { L"Midterm", 70 },
        { L"Final", 70 },
        { L"Course Avg", 80 },
        { L"Grade", 50 }
    };

    for (int i = 0; i < sizeof(columns) / sizeof(columns[0]); i++) {
        lvc.iSubItem = i;
        lvc.pszText = const_cast<LPWSTR>(columns[i].name);
        lvc.cx = columns[i].width;
        ListView_InsertColumn(hWndListView, i, &lvc);
    }
}

bool LoadStudentsFromFile(const wstring& filename) {
    students.clear();

    string filenameStr(filename.begin(), filename.end());

   ifstream file(filenameStr + ".txt");
    if (!file.is_open()) {
        return false;
    }

    string line;
    while (!file.eof()) {
        Student student;

        
        getline(file, line);
        if (line.empty() && file.eof()) break;

        student.firstName = line;

        getline(file, student.lastName);

        getline(file, line);
        student.attendance = stod(line);

        getline(file, line);
        student.groupWork = stod(line);

        
        for (int i = 0; i < 4; i++) {
            getline(file, line);
            student.quizzes[i] = stod(line);
        }

        
        for (int i = 0; i < 8; i++) {
            getline(file, line);
            student.labs[i] = stod(line);
        }

        
        for (int i = 0; i < 4; i++) {
           getline(file, line);
            student.homework[i] = stod(line);
        }

        getline(file, line);
        student.midterm = stod(line);

        getline(file, line);
        student.finalExam = stod(line);

        student.avgQuizzes = findAverageQuizScore(student.quizzes);
        student.avgLabs = findAverageLabScore(student.labs);
        student.avgHomework = findAverageHomeworkScore(student.homework);

        student.courseAverage =
            (student.attendance * 0.05) +
            (student.groupWork * 0.1) +
            (student.avgQuizzes * 0.15) +
            (student.avgLabs * 0.2) +
            (student.avgHomework * 0.2) +
            (student.midterm * 0.1) +
            (student.finalExam * 0.2);

        student.letterGrade = calculateLetterGrade(student.courseAverage);

        students.push_back(student);
    }

    file.close();
    return true;
}

void PopulateListView() {
    ListView_DeleteAllItems(hWndListView);

    for (size_t i = 0; i < students.size(); i++) {
        LVITEM lvi = { 0 };
        lvi.mask = LVIF_TEXT;
        lvi.iItem = (int)i;
        lvi.iSubItem = 0;

        
        vector<wchar_t> firstNameBuf(students[i].firstName.length() + 1);

        wstring firstNameW(students[i].firstName.begin(), students[i].firstName.end());
        wcscpy_s(&firstNameBuf[0], firstNameBuf.size(), firstNameW.c_str());
        lvi.pszText = &firstNameBuf[0];
        int index = ListView_InsertItem(hWndListView, &lvi);
        wstring lastNameW(students[i].lastName.begin(), students[i].lastName.end());
        vector<wchar_t> lastNameBuf(lastNameW.length() + 1);
        wcscpy_s(&lastNameBuf[0], lastNameBuf.size(), lastNameW.c_str());
        ListView_SetItemText(hWndListView, index, 1, &lastNameBuf[0]);

     
        wchar_t buffer[32];

        swprintf_s(buffer, L"%.2f", students[i].avgQuizzes);
        ListView_SetItemText(hWndListView, index, 2, buffer);

        swprintf_s(buffer, L"%.2f", students[i].avgLabs);
        ListView_SetItemText(hWndListView, index, 3, buffer);

        swprintf_s(buffer, L"%.2f", students[i].avgHomework);
        ListView_SetItemText(hWndListView, index, 4, buffer);

        swprintf_s(buffer, L"%.2f", students[i].midterm);
        ListView_SetItemText(hWndListView, index, 5, buffer);

        swprintf_s(buffer, L"%.2f", students[i].finalExam);
        ListView_SetItemText(hWndListView, index, 6, buffer);

        swprintf_s(buffer, L"%.2f", students[i].courseAverage);
        ListView_SetItemText(hWndListView, index, 7, buffer);

        
        wstring letterGradeW(students[i].letterGrade.begin(), students[i].letterGrade.end());
        vector<wchar_t> letterGradeBuf(letterGradeW.length() + 1);
        wcscpy_s(&letterGradeBuf[0], letterGradeBuf.size(), letterGradeW.c_str());
        ListView_SetItemText(hWndListView, index, 8, &letterGradeBuf[0]);
    }
}

bool ExportResultsToFile(const wstring& filename) {
    string filenameStr(filename.begin(), filename.end());

    ofstream file(filenameStr);
    if (!file.is_open()) {
        return false;
    }

    file << "Student Grade Results" << endl;
    file << "====================" << endl << endl;

    for (const auto& student : students) {
        file << "Name: " << student.firstName << " " << student.lastName << endl;
        file << "Attendance: " << fixed << setprecision(2) << student.attendance << endl;
        file << "Group Work: " << student.groupWork << endl;
        file << "Quiz Average: " << student.avgQuizzes << endl;
        file << "Homework Average: " << student.avgHomework << endl;
        file << "Midterm: " << student.midterm << endl;
        file << "Final Exam: " << student.finalExam << endl;
        file << "Course Average: " << student.courseAverage << endl;
        file << "Letter Grade: " << student.letterGrade << endl;
        file << "--------------------------------------------------" << endl << endl;
    }

    file.close();
    return true;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize common controls
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
    InitCommonControlsEx(&icex);

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_GRADECALCULATOR, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance(hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_GRADECALCULATOR));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_GRADECALCULATOR));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_GRADECALCULATOR);
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance;

    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, 800, 600, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd)
    {
        return FALSE;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
    {
        CreateWindowEx(0, L"STATIC", L"Enter file name (without extension):",
            WS_CHILD | WS_VISIBLE,
            10, 10, 220, 20, hWnd, NULL, hInst, NULL);

        hWndEditFilename = CreateWindowEx(
            WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            10, 30, 220, 20, hWnd, (HMENU)ID_EDIT_FILENAME, hInst, NULL);

        hWndButtonLoad = CreateWindow(
            L"BUTTON", L"Load Data",
            WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            240, 30, 100, 20, hWnd, (HMENU)ID_BUTTON_LOAD, hInst, NULL);

        hWndButtonExport = CreateWindow(
            L"BUTTON", L"Export Results",
            WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            350, 30, 100, 20, hWnd, (HMENU)ID_BUTTON_EXPORT, hInst, NULL);

        InitListView(hWnd);

        hWndStatus = CreateWindowEx(
            0, STATUSCLASSNAME, L"Ready",
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
            0, 0, 0, 0, hWnd, (HMENU)ID_STATUSBAR, hInst, NULL);
    }
    break;

    case WM_SIZE:
    {
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);

        MoveWindow(hWndListView,
            0, 60,
            rcClient.right, rcClient.bottom - 80,
            TRUE);

  
        SendMessage(hWndStatus, WM_SIZE, 0, 0);
    }
    break;

    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case ID_BUTTON_LOAD:
        {
            wchar_t filename[MAX_PATH];
            GetWindowText(hWndEditFilename, filename, MAX_PATH);

            if (wcslen(filename) == 0) {
                MessageBox(hWnd, L"Please enter a filename", L"Error", MB_ICONERROR | MB_OK);
                break;
            }

            if (LoadStudentsFromFile(filename)) {
                PopulateListView();

                wchar_t statusText[256];
                swprintf_s(statusText, L"Loaded %d students from %s.txt", (int)students.size(), filename);
                SetWindowText(hWndStatus, statusText);
            }
            else {
                MessageBox(hWnd, L"Failed to load file. Please check the filename and try again.",
                    L"File Error", MB_ICONERROR | MB_OK);
            }
        }
        break;

        case ID_BUTTON_EXPORT:
        {
            if (students.empty()) {
                MessageBox(hWnd, L"No data to export. Please load student data first.",
                    L"Warning", MB_ICONWARNING | MB_OK);
                break;
            }

            OPENFILENAME ofn;
            wchar_t szFileName[MAX_PATH] = L"GradeResults.txt";

            ZeroMemory(&ofn, sizeof(ofn));
            ofn.lStructSize = sizeof(ofn);
            ofn.hwndOwner = hWnd;
            ofn.lpstrFilter = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
            ofn.lpstrFile = szFileName;
            ofn.nMaxFile = MAX_PATH;
            ofn.Flags = OFN_EXPLORER | OFN_OVERWRITEPROMPT;
            ofn.lpstrDefExt = L"txt";

            if (GetSaveFileName(&ofn)) {
                if (ExportResultsToFile(ofn.lpstrFile)) {
                    wchar_t statusText[256];
                    swprintf_s(statusText, L"Results exported to %s", szFileName);
                    SetWindowText(hWndStatus, statusText);
                }
                else {
                    MessageBox(hWnd, L"Failed to export results", L"Error", MB_ICONERROR | MB_OK);
                }
            }
        }
        break;

        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;

        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;

        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        EndPaint(hWnd, &ps);
    }
    break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}