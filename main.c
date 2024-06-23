#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>

void gotoXY(int x, int y) {
    COORD coord;
    coord.X = y;
    coord.Y = x;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coord);
}

//다음 셀을 각각 오른쪽과 아래로 연결
typedef struct Cell {
    char data[20]; //셀의 값은 최대 19바이트
    struct Cell* right;
    struct Cell* down;
} Cell;

typedef struct {
    Cell* head;
    int rows;
    int cols;
} Spreadsheet;

//
Cell* makeCell() {
    Cell* head = (Cell*)malloc(sizeof(Cell));
    head->data[0] = '\0';
    head->right = NULL;
    head->down = NULL;
    return head;
}

Spreadsheet* makeSpreadsheet(int rows, int cols) {
    Spreadsheet* sheet = (Spreadsheet*)malloc(sizeof(Spreadsheet));
    sheet->rows = rows;
    sheet->cols = cols;
    sheet->head = makeCell();

    Cell* current = sheet->head;
    for (int r = 0; r < rows; ++r) {
        Cell* rowHead = current; 
        for (int c = 1; c < cols; ++c) { //셀을 오른쪽에 만들고 연결하여 열을 만든다
            current->right = makeCell();
            current = current->right;
        }
        if (r < rows - 1) { //셀을 아래쪽에 만들고 연결하여 행을 만든다
            rowHead->down = makeCell();
            current = rowHead->down;
        }
    }
    return sheet;
}

void freeSpreadsheet(Spreadsheet* sheet) {
    Cell* row = sheet->head;
    while (row != NULL) {
        Cell* col = row;
        while (col != NULL) {
            Cell* tempCol = col;
            col = col->right;
            free(tempCol);
        }
        Cell* tempRow = row;
        row = row->down;
        free(tempRow);
    }
    free(sheet);
}

void printSpreadsheet(Spreadsheet* sheet) {
    gotoXY(0, 0);
    printf("    ");
    for (int col = 0; col < sheet->cols; ++col) { // A~AZ 총 52열까지 열 번호 표시
        if ('A' + col > 'Z') {
            printf("  A%c ", 'A' + col - 26);
        }
        else {
            printf("  %c  ", 'A' + col);
        }
    }
    printf("\n");

    Cell* row = sheet->head;
    for (int r = 0; r < sheet->rows; ++r) { // 행 번호 표시
        printf("%2d ", r + 1);
        Cell* col = row;
        for (int c = 0; c < sheet->cols; ++c) {
            printf(" %3s ", col->data[0] ? col->data : " "); // 빈 셀은 공백표시
            col = col->right;
        }
        printf("\n");
        row = row->down;
    }
}

void Range(const char* range, int* startRow, int* startCol, int* endRow, int* endCol) { //범위지정
    char startCell[5], endCell[5];
    sscanf(range, "%4[^-]-%4s", startCell, endCell);

    *startRow = atoi(&startCell[1]) - 1; 
    *startCol = startCell[0] - 'A';
    *endRow = atoi(&endCell[1]) - 1;
    *endCol = endCell[0] - 'A';
    //A1-C3 입력이면 startRow=0, startCol=0, endRow=2, endCol=2
}

void setValue(Spreadsheet* sheet, int row, int col, const char* value) { //셀 값 설정
    if (row >= 0 && row < sheet->rows && col >= 0 && col < sheet->cols) {
        Cell* current = sheet->head;
        for (int r = 0; r < row; ++r) {
            current = current->down;
        }//행 이동
        for (int c = 0; c < col; ++c) {
            current = current->right;
        }//열 이동
        strncpy(current->data, value, 19);
        current->data[19] = '\0'; 
    }
}

char* getValue(Spreadsheet* sheet, int row, int col) { //셀 값 불러오기
    if (row >= 0 && row < sheet->rows && col >= 0 && col < sheet->cols) {
        Cell* current = sheet->head;
        for (int r = 0; r < row; ++r) {
            current = current->down;
        }
        for (int c = 0; c < col; ++c) {
            current = current->right;
        }
        return current->data;
    }
}

void add(Spreadsheet* sheet, int num, int rowcol) { //행, 열 추가
    if (rowcol == 0) { //addrow 입력 시 rowcol==0
        for (int i = 0; i < num; ++i) {
            Cell* current = sheet->head;
            while (current->down != NULL) { 
                current = current->down;
            }
            current->down = makeCell(); //맨 아래에 행 추가
            Cell* newRow = current->down;
            for (int c = 1; c < sheet->cols; ++c) {
                newRow->right = makeCell();
                newRow = newRow->right;
            }//새로운 행에 셀 추가
        }
        sheet->rows += num;
    }
    else {
        for (int r = 0; r < sheet->rows; ++r) {
            Cell* current = sheet->head;
            for (int i = 0; i < r; ++i) { // 행 이동
                current = current->down;
            }
            while (current->right != NULL) { // 행의 마지막 열로 이동
                current = current->right;
            }
            for (int i = 0; i < num; ++i) { //셀 추가
                current->right = makeCell();
                current = current->right;
            }
        }
        sheet->cols += num;
    }

}

void processCommand(Spreadsheet* sheet, const char* command) {
    if (strncmp(command, "set", 3) == 0) {
        char input[50];
        strcpy(input, command + 4); //set A1-A3=10  A부터 input으로 복사
        char* token = strtok(input, "="); 
        char range[10];
        char value[20];
        sscanf(token, "%9s", range); //range에 A1-A3 저장
        token = strtok(NULL, "="); 
        if (token) {
            sscanf(token, "%19s", value); //value에 '=' 뒷부분 저장(값)
            int startRow, startCol, endRow, endCol;
            if (strchr(range, '-') == NULL) { //단일 셀 입력
                startCol = range[0] - 'A';
                startRow = atoi(&range[1]) - 1;
                setValue(sheet, startRow, startCol, value);
            }
            else { //범위 입력
                Range(range, &startRow, &startCol, &endRow, &endCol);
                for (int row = startRow; row <= endRow; ++row) {
                    for (int col = startCol; col <= endCol; ++col) {
                        setValue(sheet, row, col, value);
                    }
                }
            }
        }
    }
    else if (strncmp(command, "get", 3) == 0) {
        char input[50];
        strcpy(input, command + 4);
        if (strchr(input, '-') == NULL) { // 단일 셀
            int row = atoi(&input[1]) - 1;
            int col = input[0] - 'A';
            printf("%c%d: %s\n", input[0], row + 1, getValue(sheet, row, col));
        }
        else { // 범위
            int startRow, startCol, endRow, endCol;
            Range(input, &startRow, &startCol, &endRow, &endCol);
            for (int row = startRow; row <= endRow; ++row) {
                for (int col = startCol; col <= endCol; ++col) {
                    printf("%c%d: %s\n", 'A' + col, row + 1, getValue(sheet, row, col));
                }
            }
        }
        system("pause"); 
    }
    else if (strncmp(command, "addrow", 6) == 0 || strncmp(command, "addcol", 6) == 0) {
        int num = atoi(command + 7);
        add(sheet, num, strncmp(command, "addrow", 6));
    }

    else if (strncmp(command, "sum", 3) == 0) {
        char input[50];
        strcpy(input, command + 4);
        
        if (strchr(input, '+') != NULL) {
            char cell1[5], cell2[5];
            sscanf(input, "%4[^+]+%4s", cell1, cell2);
            int row1 = atoi(&cell1[1]) - 1;
            int col1 = cell1[0] - 'A';
            int row2 = atoi(&cell2[1]) - 1;
            int col2 = cell2[0] - 'A';
            int value1 = atoi(getValue(sheet, row1, col1));
            int value2 = atoi(getValue(sheet, row2, col2));
            printf("%d\n", value1 + value2);
        }
        else {
            int startRow, startCol, endRow, endCol, sum = 0;
            Range(input, &startRow, &startCol, &endRow, &endCol);
            for (int row = startRow; row <= endRow; ++row) {
                for (int col = startCol; col <= endCol; ++col) {
                    sum += atoi(getValue(sheet, row, col));
                }
            }
            printf("%d\n", sum);
        }      
        system("pause");
    }
    else if (strncmp(command, "sub", 3) == 0) {
        char input[50];
        strcpy(input, command + 4);
        char cell1[5], cell2[5];
        sscanf(input, "%4[^-]-%4s", cell1, cell2);
        int row1 = atoi(&cell1[1]) - 1;
        int col1 = cell1[0] - 'A';
        int row2 = atoi(&cell2[1]) - 1;
        int col2 = cell2[0] - 'A';
        int value1 = atoi(getValue(sheet, row1, col1));
        int value2 = atoi(getValue(sheet, row2, col2));
        printf("%d\n", value1 - value2);
        system("pause");
    }
    else if (strncmp(command, "product", 7) == 0) {
        char input[50];
        strcpy(input, command + 8);
        if (strchr(input, '*') != NULL) {
            char cell1[5], cell2[5];
            sscanf(input, "%4[^*]*%4s", cell1, cell2);
            int row1 = atoi(&cell1[1]) - 1;
            int col1 = cell1[0] - 'A';
            int row2 = atoi(&cell2[1]) - 1;
            int col2 = cell2[0] - 'A';
            int value1 = atoi(getValue(sheet, row1, col1));
            int value2 = atoi(getValue(sheet, row2, col2));
            printf("%d\n", value1 * value2);
        }
        else {
            int startRow, startCol, endRow, endCol, prod = 1;
            Range(input, &startRow, &startCol, &endRow, &endCol);
            for (int row = startRow; row <= endRow; ++row) {
                for (int col = startCol; col <= endCol; ++col) {
                    prod *= atoi(getValue(sheet, row, col));
                }
            }
            printf("%d\n", prod);
        }        
        system("pause");
    }
    else if (strncmp(command, "div", 3) == 0) {
        char input[50];
        strcpy(input, command + 4);
        char cell1[5], cell2[5];
        sscanf(input, "%4[^/]/%4s", cell1, cell2);
        int row1 = atoi(&cell1[1]) - 1;
        int col1 = cell1[0] - 'A';
        int row2 = atoi(&cell2[1]) - 1;
        int col2 = cell2[0] - 'A';
        int value1 = atoi(getValue(sheet, row1, col1));
        int value2 = atoi(getValue(sheet, row2, col2));
        if (value2 != 0) {
            printf("%d\n", value1 / value2);
        }
        else {
            printf("Error: Division by zero.\n");
        }
        system("pause");
    }
    else if (strncmp(command, "graph", 5) == 0) {
        char input[50];
        strcpy(input, command + 6);
        int startRow, startCol, endRow, endCol, y = 0;
        Range(input, &startRow, &startCol, &endRow, &endCol);

        int maxStars = 50; // 최대 별 개수
        int maxValue = 0;
        for (int row = startRow; row <= endRow; ++row) {
            for (int col = startCol; col <= endCol; ++col) {
                int value = atoi(getValue(sheet, row, col));
                if (value > maxValue) {
                    maxValue = value; // 그래프가 스케일 조절을 위한 최댓값
                }
            }
        }
        int scale = maxValue / maxStars;
        if (scale == 0) scale = 1;

        printf("\n");
        for (int row = startRow; row <= endRow; ++row) {
            for (int col = startCol; col <= endCol; ++col) {
                int value = atoi(getValue(sheet, row, col));
                int stars = value / scale;
                printf("%c%d | %3d ", 'A' + col, row + 1, value);
                for (int i = 0; i < stars; i++) {
                    printf("*");
                }
                printf("\n");
            }
        }
        system("pause");
    }
    else {
        printf("다시 입력..\n");
    }
}

void run(Spreadsheet* sheet) {
    char command[50];
    while (1) {
        system("cls"); 
        printSpreadsheet(sheet);
        printf("명령어 입력(exit 입력 시 종료): ");
        printf("\n1.set 2.get 3.addrow/col 4.sum 5.product 6.div 7.graph\n");
        if (fgets(command, sizeof(command), stdin) == NULL) {
            printf("Error reading input\n");
            continue;
        }
        command[strcspn(command, "\n")] = 0; 
        if (strcmp(command, "exit") == 0) break;
        processCommand(sheet, command);
    }
}

int main() {
    int rows, cols;
    printf("스프레드시트 크기 입력(행, 열): ");
    scanf("%d %d", &rows, &cols);
    getchar();

    Spreadsheet* sheet = makeSpreadsheet(rows, cols);
    run(sheet);
    freeSpreadsheet(sheet);

    return 0;
}
