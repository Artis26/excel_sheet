<?php
namespace App\Models;

use Carbon\Carbon;
use Carbon\CarbonPeriod;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class Sheet {
    public function getLastFourMonths($year, $month): array {
        for ($i = 1; $i < 5; $i++) {
            $date = Carbon::parse("$year-$month-01");
            $date = $date->subMonths($i);
            $lastFourMonths[] = $date->monthName;
        }
        return $lastFourMonths;
    }

    public function generateDateCells($year, $month, Spreadsheet $spreadsheet) {
        $sheet = $spreadsheet->getActiveSheet();
        $columnLocation = 3;
        $firstDay = "$year-$month-01";
        $lastDayOfMonth = Carbon::parse($firstDay)->endOfMonth()->toDateString();
        $period = CarbonPeriod::create($firstDay, $lastDayOfMonth);

        foreach ($period as $key => $date) {
            $sheet->setCellValueByColumnAndRow($columnLocation, 8, $key + 1)->getColumnDimensionByColumn($columnLocation)->setWidth(25, 'px');
            if ($date->isWeekend() == 1) {
                $sheet->getStyleByColumnAndRow($columnLocation, 8)
                    ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB(Color::COLOR_CYAN);
            }
            $columnLocation++;
        }
    }

    public function generatePersonCells(Spreadsheet $spreadsheet, array $persons) {
        $sheet = $spreadsheet->getActiveSheet();
        $sizeOfOnePersonInCells = 3;
        $start = 9;
        $end = 11;
        foreach ($persons as $key => $person) {
            $sheet->mergeCellsByColumnAndRow(1, $start, 1, $end)->setCellValueByColumnAndRow(1, $start, $key + 1);
            $sheet->getStyleByColumnAndRow(1, $start)->getAlignment()->setVertical('center')->setHorizontal('center');

            $richText = new RichText();
            $payable = $richText->createTextRun($person->getName());
            $payable->getFont()->setBold(true);
            $richText->createText("\n" . $person->getPosition());

            $sheet->setCellValueByColumnAndRow(2, $start, $richText);
            $sheet->getStyleByColumnAndRow(2, $start)->getAlignment()->setWrapText(true);

            $sheet->setCellValueByColumnAndRow(2, $start + 1, "t.sk. 22.00-06.00");
            $sheet->setCellValueByColumnAndRow(2, $end, "izsaukums (plkst. no l??dz)")->getStyleByColumnAndRow(2, $end)->getFont()->setSize(8);

            $sheet->getStyleByColumnAndRow(1, $start, 54, $end)->getBorders()->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN)->setColor(new Color('#000000'));
            $sheet->getStyleByColumnAndRow(1, $start, 33, $end)->getBorders()->getOutline()
                ->setBorderStyle(Border::BORDER_THICK)->setColor(new Color('#000000'));
            $sheet->getStyleByColumnAndRow(34, $start, 54, $end)->getBorders()->getOutline()
                ->setBorderStyle(Border::BORDER_THICK)->setColor(new Color('#000000'));
            $sheet->getStyleByColumnAndRow(3, $start, 3, $end)->getBorders()->getLeft()
                ->setBorderStyle(Border::BORDER_THICK)->setColor(new Color('#000000'));

            $sheet->getStyle("C$start:BB$end")->getAlignment()->setHorizontal('center')->setVertical('center');

            //Adding cell function
            $sum = "C$start:AG$start";
            $sheet->setCellValueByColumnAndRow(35, $start, "=SUM($sum)");

            $middle = $start + 1;
            $nightSum = "C$middle:AG$middle";
            $sheet->setCellValueByColumnAndRow(36, $middle, "=SUM($nightSum)");

            $start += $sizeOfOnePersonInCells;
            $end += $sizeOfOnePersonInCells;
        }
    }

    public function generateOvertimeCells(Spreadsheet $spreadsheet, $columnWhereStart) {
        $sheet = $spreadsheet->getActiveSheet();
        $column = $columnWhereStart;
        $rowStart = 6;
        $rowEnd = 8;
        for($i=2; $i<7; $i++) {
            $sheet->mergeCellsByColumnAndRow($column, $rowStart, $column, $rowEnd)->setCellValueByColumnAndRow($column, $rowStart, "virs noteikt?? dienesta pien??kumu izpildes laika" . $i)
                ->getStyleByColumnAndRow($column, $rowStart)->getAlignment()->setTextRotation(90);

            $column++;
        }
    }

    public function hideEmptyCells(Spreadsheet $spreadsheet) {
        $sheet = $spreadsheet->getActiveSheet();
        $start = 30;
        $end = 33;
        $dateRow = 8;
        for($i=0; $i<4; $i++) {
            if (empty($sheet->getCellByColumnAndRow($start, $dateRow)->getValue())) {
                $sheet->getColumnDimensionByColumn($start)->setWidth(0);
            }
            $start++;
        }
    }

    public function generateNew($year, $month) {
        $allMonths = ['janv??ri', 'febru??ri', 'martu', 'apr??li', 'maiju', 'j??niju', 'j??liju', 'augustu', 'septembri', 'oktobri', 'novembri', 'decembri'];
        setlocale(LC_TIME, 'lv_LV');
        Carbon::setLocale('lv');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $persons = [
            new Person('J??nis', 'pav??rs'),
            new Person('Andris', 'apsargs'),
            new Person('Aivars', 'vec??kais inspektors'),
            new Person('Aldis', 'meh??ni??is')
        ];

        $sheet->mergeCells("C2:AD2")->setCellValue("C2", "DIENESTA PIEN??KUMU IZPILDES(DARBA) LAIKA UZSKAITES TABULA")
            ->getStyle("C2")->getFont()->setBold(true);
        $sheet->getStyle("C2")->getAlignment()->setHorizontal('center')->setVertical('center');

        $sheet->mergeCells("C3:AD3")->setCellValue("C3", "par $year.gada " . $allMonths[$month-1])
            ->getStyle("C3:AD3")->getFont()->setBold(true);
        $sheet->getStyle("C3:AD3")->getAlignment()->setHorizontal('center')->setVertical('center');

        $sheet->mergeCells("A5:A8")->setCellValue("A5", "nr.p/k")
            ->getStyle("A5")->getAlignment()->setTextRotation(90)->setVertical('center');
        $sheet->getColumnDimension("A")->setWidth(20, 'px');

        $sheet->mergeCells("C5:AG7")->setCellValue("C5", "M??ne??a datumi")
            ->getStyle("C5:AG7")->getAlignment()->setVertical('center');

        $sheet->mergeCells("B5:B8")->setCellValue("B5", "v??rds, uzv??rds, amats")
            ->getStyle("B5:B8")->getAlignment()->setVertical('center')->setWrapText(true);

        $this->generateDateCells($year, $month, $spreadsheet);

        $sheet->mergeCells("AH5:AH8")->setCellValue("AH5", "stundu skaits m??ne??a norm??laj?? dien.pien??kumu izpildes (darba) laik??")
            ->getStyle("AH5")->getAlignment()->setTextRotation(90)->setWrapText(true);

        $sheet->mergeCells("AI5:AW5")->setCellValue("AI5", "Nostr??d??t??s stundas")->getStyle("AI5")->getFont()->setBold(true);
        $sheet->getStyle("AI5")->getAlignment()->setVertical('center');

        $sheet->mergeCells("AI6:AI8")->setCellValue("AI6", "pavisam")
            ->getStyle("AI6")->getAlignment()->setTextRotation(90);
        $sheet->getStyle("AI6")->getFont()->setBold(true);

        $sheet->mergeCells("AJ6:AJ8")->setCellValue("AJ6", "nakt??(no plkst.22.00 l??dz 6.00")
            ->getStyle("AJ6")->getAlignment()->setTextRotation(90);

        $sheet->mergeCells("AK6:AK8")->setCellValue("AK6", "de????ras ??rpus dienesta pien??kumu izpildes vietas")
            ->getStyle("AK6")->getAlignment()->setTextRotation(90);

        $sheet->mergeCells("AL6:AL8")->setCellValue("AL6", "virsstundas")
            ->getStyle("AL6")->getAlignment()->setTextRotation(90);

        $this->generateOvertimeCells($spreadsheet, 39);

        $sheet->mergeCells("AR6:AR8")->setCellValue("AR6", "sv??tku dienas (darbiniekiem)")
            ->getStyle("AR6")->getAlignment()->setTextRotation(90);

        $sheet->mergeCells("AS6:AW7")->setCellValue("AS6", "virs noteikt?? dienesta pien??kumu izpildes laika (virstundas)")
            ->getStyle("AS6")->getAlignment()->setWrapText(true);
        $sheet->getStyle("As6")->getFont()->setSize(8);

        $sheet->setCellValue("AS8", "kop?? iepriek????jo 4 m??ne??u period??")
            ->getStyle("AS8")->getAlignment()->setTextRotation(90);

        $months = $this->getLastFourMonths($year, $month);
        $sheet->setCellValue("AT8", "pirmaj?? m??nes??($months[3])")
            ->getStyle("AT8")->getAlignment()->setTextRotation(90);

        $sheet->setCellValue("AU8", "otraj?? m??nes??($months[2])")
            ->getStyle("AU8")->getAlignment()->setTextRotation(90);

        $sheet->setCellValue("AV8", "tre??aj?? m??nes??($months[1])")
            ->getStyle("AV8")->getAlignment()->setTextRotation(90);

        $sheet->setCellValue("AW8", "ceturtaj?? m??nes??($months[0])")
            ->getStyle("AW8")->getAlignment()->setTextRotation(90);

        $sheet->mergeCells("AX5:AX8")->setCellValue("AX5", "virs noteikt?? dienesta pien??k. izp. laika kop?? ??etru m??ne??u period??")
            ->getStyle("AX5")->getAlignment()->setTextRotation(90)->setWrapText(true);

        $sheet->mergeCells("AY5:AY8")->setCellValue("AY5", "darba dien??s")
            ->getStyle("AY5")->getAlignment()->setTextRotation(90);
        $sheet->getStyle("AY5")->getFont()->setBold(true);

        $sheet->mergeCells("AZ5:AZ8")->setCellValue("AZ5", "atva??in??juma dienas")
            ->getStyle("AZ5")->getAlignment()->setTextRotation(90);
        $sheet->getStyle("AZ5")->getFont()->setBold(true);

        $sheet->mergeCells("BA5:BA8")->setCellValue("BA5", "darbnesp??jas dienas")
            ->getStyle("BA5")->getAlignment()->setTextRotation(90);
        $sheet->getStyle("BA5")->getFont()->setBold(true);

        $sheet->mergeCells("BB5:BB8")->setCellValue("BB5", "apmaks??ts atp??tas laiks")
            ->getStyle("BB5")->getAlignment()->setTextRotation(90);
        $sheet->getStyle("BB5")->getFont()->setBold(true);

        $this->generatePersonCells($spreadsheet, $persons);

        //Style table part without persons:
        $sheet->getColumnDimension("B")->setWidth(125, 'px');
        $sheet->getStyle("A5:BB8")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN)->setColor(new Color('#000000'));
        $sheet->getStyle("A5:BB8")->getAlignment()->setHorizontal('center');
        $sheet->getRowDimension('8')->setRowHeight(325, 'px');
        foreach (range(34, 54) as $columnID) {
            $sheet->getColumnDimensionByColumn($columnID)->setWidth(30, 'px');
        }
        $sheet->getStyle("C8")->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK)->setColor(new Color('#000000'));
        $this->hideEmptyCells($spreadsheet);

        return $spreadsheet;
    }
}
