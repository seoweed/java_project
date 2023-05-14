package kr.excel.example;

import kr.excel.entity.MemberVO;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelWriter {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        List<MemberVO> members = new ArrayList<>();
        while (true) {
            System.out.print("이름을 입력하세요 : ");
            String name = scanner.nextLine();
            if (name.equals("quit")){
                break;
            }

            System.out.print("나이를 입력하세요 : ");
            int age = scanner.nextInt();
            scanner.nextLine();

            System.out.print("생년월일을 입력하세요 : ");
            String birthday = scanner.nextLine();

            System.out.print("전화번호를 입력하세요 : ");
            String phone = scanner.nextLine();

            System.out.print("주소를 입력하세요 : ");
            String address = scanner.nextLine();

            System.out.print("결혼여부를 입력하세요 : ");
            boolean isMarried = scanner.nextBoolean();
            scanner.nextLine();

            MemberVO memberVO = new MemberVO(name, age, birthday,phone, address, isMarried);
            members.add(memberVO);
        }
        scanner.close();

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("회원 정보");

            XSSFRow headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("이름");
            headerRow.createCell(1).setCellValue("나이");
            headerRow.createCell(2).setCellValue("생년월일");
            headerRow.createCell(3).setCellValue("전화번호");
            headerRow.createCell(4).setCellValue("주소");
            headerRow.createCell(5).setCellValue("결혼여부");

            for (int i = 0; i < members.size(); i++) {
                MemberVO member = members.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(member.getName());
                row.createCell(1).setCellValue(member.getAge());
                row.createCell(2).setCellValue(member.getBirthday());
                row.createCell(3).setCellValue(member.getPhone());
                row.createCell(4).setCellValue(member.getAddress());
                Cell marriedCell = row.createCell(5);
                marriedCell.setCellValue(member.isMarried());

            }
            String filename = "members.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(new File(filename));
            workbook.write(fileOutputStream);
            workbook.close();
            System.out.println("엑셀파일이 저장되었습니다." + filename);

        }catch (IOException e) {
            System.out.println("엑셀 파일 저장 중 오류가 발생하였습니다");
            e.printStackTrace();
        }

    }
}
