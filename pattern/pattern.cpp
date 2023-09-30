#include <OpenXLSX.hpp>
#include <string>
#include <array>
using namespace OpenXLSX;

#define PATTERN_FILE "pattern.xlsx"
#define SHEET "Sheet1"

// какие файлы открывать
std::array<std::string, 4> openName = {"petrov", "denisov", "krayn", "vertalov"};

int main()
{
    // куда записывать
    XLDocument pattern;
    pattern.open(PATTERN_FILE);
    auto pattern_sheet = pattern.workbook().worksheet(SHEET);

    XLDocument p;
    for (int i = 0; i < openName.size(); i++)
    {
        // плюс два, счет 1 и первая строка занята заголовком
        std::string index = std::to_string(i + 2);

        // открываем досье
        p.open(openName[i] + ".xlsx");
        auto p_sheet = p.workbook().worksheet(SHEET);

        // заполняем шаблон

        // имя
        pattern_sheet.cell("A" + index).value() = p_sheet.cell("A1").value();

        // фамилия
        pattern_sheet.cell("B" + index).value() = p_sheet.cell("A2").value();

        // возраст
        pattern_sheet.cell("C" + index).value() = p_sheet.cell("A3").value();

        // улица
        pattern_sheet.cell("D" + index).value() = p_sheet.cell("A4").value();
    }

    pattern.save();

    return 0;
}