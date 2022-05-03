<?php

namespace App\Http\Controllers;

use App\Models\Sheet;
use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Routing\Controller as BaseController;
use Illuminate\Support\Facades\Response;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Controller extends BaseController {
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    public function generate() {
        $newSheet = new Sheet();
        $newSheet = $newSheet->generateNew(2022, 2);

        $writer = new xlsx($newSheet);
        $writer->save("demo.xlsx");
        echo "Generating .xlsx file";
    }

   public function export() {
       $filepath = public_path('demo.xlsx');
       return Response::download($filepath);
   }
}
