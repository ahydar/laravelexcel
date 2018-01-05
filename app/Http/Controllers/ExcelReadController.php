<?php

namespace App\Http\Controllers;

use App\ExcelReader\PhpExcelReader;
use Illuminate\Http\Request;

class ExcelReadController extends Controller
{
    //
    public function index()
    {
        $sheet = new PhpExcelReader;

        var_dump($sheet);
        return "Created";
    }

    public function store(Request $request)
    {
        $filename = $request->file('marksheet')->getRealPath();
        
        $excel = new PhpExcelReader;

        $excel->read($filename);
        
        $sheet = $excel->sheets[0];

        $marks = [];

        for($i=0;$i<$sheet['numRows']-12;$i++){
            if(isset($sheet['cells'][12+$i][2])){
                $data = [
                    'Name' => $sheet['cells'][12+$i][2],
                    'mark1' => $sheet['cells'][12+$i][27],
                    'mark2' => $sheet['cells'][12+$i][28],
                    'mark3' => $sheet['cells'][12+$i][29],
                    'mark4' => $sheet['cells'][12+$i][30],
                    'final' => $sheet['cells'][12+$i][38],
                ];
            }
            
            $marks[] = $data;
        }
        $request->file('marksheet')->store('marksheets');
        return $marks;
    }
}
