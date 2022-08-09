package org.example;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.sql.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class OneClickReport {
    public static void main(String[] args) {
        Report report = new Report();
        report.initUI();
    }
}

class Report extends JFrame {
    String excelUrl1 = ""; //Excel地址
    String mdbUrl = ""; //mdb地址
    final JTextField textFieldExcel = new JTextField(50); //设置Excel文本框
    final JTextField textFieldMdb = new JTextField(50); //设置mdb文本框
    final JButton buttonDoReport = new JButton("填写报告");//设置报告填写按钮
    final JRadioButton radioButtonCN = new JRadioButton("中文", true);//中文选项
    final JRadioButton radioButtonEN = new JRadioButton("英文", false);//英文选项

    //创建界面
    public void initUI() {
        this.setTitle("填报告-20210515"); //设置标题
        this.setSize(600, 180); //设置大小
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setResizable(false); //设置不可拉伸
        this.setLayout(null); //关闭流式布局
        //开始界面元素设计
        final JLabel labelExcelDir = new JLabel("Excel目录"); //设置目录字样
        labelExcelDir.setBounds(50, 5, 60, 25); //设置位置大小
        this.add(labelExcelDir); //添加标签
        final JLabel labelMdbDir = new JLabel("数据库Data目录"); //设置目录字样
        labelMdbDir.setBounds(13, 35, 100, 25); //设置位置大小
        this.add(labelMdbDir); //添加标签

        textFieldExcel.setBounds(113, 5, 280, 25); //设置位置大小
        textFieldExcel.setEditable(false);//设置不可编辑
        textFieldExcel.setBackground(Color.white);//设置文本框背景色为白色
        textFieldExcel.setText("");
        this.add(textFieldExcel); //添加标签
        textFieldMdb.setBounds(113, 37, 280, 25); //设置位置大小
        textFieldMdb.setEditable(false);//设置不可编辑
        textFieldMdb.setBackground(Color.white);//设置文本框背景色为白色
        textFieldMdb.setText("");
        this.add(textFieldMdb); //添加标签

        final JButton buttonExcelDir = new JButton("获取Excel报告位置"); //设置目录按钮
        buttonExcelDir.setBounds(400, 5, 150, 25); //设置位置大小
        this.add(buttonExcelDir); //添加按钮
        final JButton buttonMdbDir = new JButton("获取数据库Data位置"); //设置目录按钮
        buttonMdbDir.setBounds(400, 37, 170, 25); //设置位置大小
        this.add(buttonMdbDir); //添加按钮
        buttonDoReport.setBounds(400, 80, 100, 50);
        this.add(buttonDoReport);
        radioButtonCN.setBounds(120, 65, 60, 60);
        this.add(radioButtonCN);
        radioButtonEN.setBounds(200, 65, 60, 60);
        this.add(radioButtonEN);
        final ButtonGroup buttonGroup = new ButtonGroup();
        buttonGroup.add(radioButtonCN);
        buttonGroup.add(radioButtonEN);
        //结束界面元素设计

        ButtonListener buttonListener = new ButtonListener(this);
        buttonExcelDir.addActionListener(buttonListener);
        buttonMdbDir.addActionListener(buttonListener);
        buttonDoReport.addActionListener(buttonListener);

        this.setVisible(true); //设置窗体可见

    }

    //打开Excel文件对话框
    public void showFileOpenDialog() {
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File("."));//设置默认显示为当前文件夹
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);//设置选择模式（文件和文件夹都可选）
        chooser.setMultiSelectionEnabled(false);//是否允许多选
        chooser.addChoosableFileFilter(new FileNameExtensionFilter("Excel(*.xls)", "xls"));//文件过滤器
        int result = chooser.showOpenDialog(null); //打开对话框
        if (result == JFileChooser.APPROVE_OPTION) {
            excelUrl1 = chooser.getSelectedFile().getAbsolutePath();
            System.out.println(excelUrl1.replaceAll("\\\\", "\\\\\\\\") + " -> showFileOpenDialog()");
            textFieldExcel.setText(excelUrl1);
        }
    }

    //打开DB文件对话框
    public void showDBOpenDialog() {
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File("."));//设置默认显示为当前文件夹
        chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);//设置选择模式（文件和文件夹都可选）
        chooser.setMultiSelectionEnabled(false);//是否允许多选
        chooser.addChoosableFileFilter(new FileNameExtensionFilter("Access(*.mdb)", "mdb"));//文件过滤器
        int result = chooser.showOpenDialog(null); //打开对话框
        if (result == JFileChooser.APPROVE_OPTION) {
//            mdbUrl = chooser.getSelectedFile().getPath();
            mdbUrl = chooser.getSelectedFile().getAbsolutePath();
            System.out.println(mdbUrl.replaceAll("\\\\", "\\\\\\\\") + " -> showDBOpenDialog()");
            textFieldMdb.setText(mdbUrl);
        }
    }

    //执行填写
    public void showExcel() throws SQLException {
        if (!excelUrl1.equals("") && !mdbUrl.equals("")) {
            if (radioButtonCN.isSelected()) {
                try {
                    cn1();
                } catch (Exception e) {
                    e.printStackTrace();
                    //JOptionPane.showMessageDialog(null, "有错误", null, JOptionPane.ERROR_MESSAGE);
                }
            }
            if (radioButtonEN.isSelected()) {
                try {
                    String excelUrl = excelUrl1.replaceAll("\\\\", "\\\\\\\\"); //将地址转义
                    POIFSFileSystem fileSystem = new POIFSFileSystem(new BufferedInputStream(new FileInputStream(excelUrl))); //使用POI读取Excel
                    HSSFWorkbook workbook = new HSSFWorkbook(fileSystem); //创建HSSFWorkbook对象
                    HSSFSheet sheet = workbook.getSheetAt(0); //读取第一个Sheet

                    Connection connection = null; //创建Connection对象
                    PreparedStatement preparedStatement = null; //创建PreparedStatement对象
                    Statement statement = null;
                    ResultSet resultSet = null; //创建ResultSet
                    String dbUrl = "jdbc:ucanaccess://" + mdbUrl.replaceAll("\\\\", "\\\\\\\\"); //创建连接
                    connection = DriverManager.getConnection(dbUrl, "username", "password"); //读取连接，在这里填写数据库密码

                    String sqlDeleteData = "delete * from TResult";

                    String sqlNo400 = "select * from TResult where tableno='no400' and nindex=30";

                    String sqlNo403 = "select * from TResult where tableno='no403' and nindex=26";

                    String sqlNo412 = "select * from TResult where tableno='no412' and nindex=26";

                    String sqlNo418 = "select * from TResult where tableno='no418' and nindex=26";

                    String sqlNo424 = "select * from TResult where tableno='no424' and nindex=26";

                    String sqlNo430 = "select * from TResult where tableno='no430' and nindex=28";

                    String sqlNo436 = "select * from TResult where tableno='no436' and porder=9";

                    /*-----初值-----*/
                    preparedStatement = connection.prepareStatement(sqlNo400);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceA22 = new CellReference("A22");
                        final HSSFRow rowA22 = sheet.getRow(cellReferenceA22.getRow());
                        final HSSFCell cellA22 = rowA22.getCell(cellReferenceA22.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellA22.setCellValue(od1);
                    }
                    /*-----初值-----*/
                    /*-----数据1-----*/
                    preparedStatement = connection.prepareStatement(sqlNo403);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceG27 = new CellReference("G27");
                        final HSSFRow rowG27 = sheet.getRow(cellReferenceG27.getRow());
                        final HSSFCell cellG27 = rowG27.getCell(cellReferenceG27.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellG27.setCellValue(od1);
                    }
                    /*-----数据2-----*/
                    preparedStatement = connection.prepareStatement(sqlNo412);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceG32 = new CellReference("G32");
                        final HSSFRow rowG32 = sheet.getRow(cellReferenceG32.getRow());
                        final HSSFCell cellG32 = rowG32.getCell(cellReferenceG32.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellG32.setCellValue(od1);
                    }
                    /*-----数据3-----*/
                    preparedStatement = connection.prepareStatement(sqlNo418);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceG37 = new CellReference("G37");
                        final HSSFRow rowG37 = sheet.getRow(cellReferenceG37.getRow());
                        final HSSFCell cellG37 = rowG37.getCell(cellReferenceG37.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellG37.setCellValue(od1);
                    }
                    /*-----数据4-----*/
                    preparedStatement = connection.prepareStatement(sqlNo424);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceG42 = new CellReference("G42");
                        final HSSFRow rowG42 = sheet.getRow(cellReferenceG42.getRow());
                        final HSSFCell cellG42 = rowG42.getCell(cellReferenceG42.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellG42.setCellValue(od1);
                    }
                    /*-----数据5-----*/
                    preparedStatement = connection.prepareStatement(sqlNo430);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceD53 = new CellReference("D53");
                        final HSSFRow rowD53 = sheet.getRow(cellReferenceD53.getRow());
                        final HSSFCell cellD53 = rowD53.getCell(cellReferenceD53.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellD53.setCellValue(od1);
                    }
                    /*-----数据5-----*/

                    /*-----数据6-----*/
                    preparedStatement = connection.prepareStatement(sqlNo436);
                    resultSet = preparedStatement.executeQuery();
                    while (resultSet.next()) {
                        final CellReference cellReferenceD61 = new CellReference("D61");
                        final HSSFRow rowD61 = sheet.getRow(cellReferenceD61.getRow());
                        final HSSFCell cellD61 = rowD61.getCell(cellReferenceD61.getCol());
                        Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                        od1 = (double) Math.round(od1 * 10000) / 10000;
                        cellD61.setCellValue(od1);
                    }
                    /*-----重新计算初值1-----*/
                    final CellReference cellReferenceP22 = new CellReference("P22");
                    final HSSFRow rowP22 = sheet.getRow(cellReferenceP22.getRow());
                    final HSSFCell cellP22 = rowP22.getCell(cellReferenceP22.getCol());
                    cellP22.setCellFormula(cellP22.getCellFormula());
                    /*-----重新计算初值1-----*/
                    /*-----重新计算数据1-----*/
                    final CellReference cellReferenceP27 = new CellReference("P27");
                    final HSSFRow rowP27 = sheet.getRow(cellReferenceP27.getRow());
                    final HSSFCell cellP27 = rowP27.getCell(cellReferenceP27.getCol());
                    cellP27.setCellFormula(cellP27.getCellFormula());
                    /*-----重新计算数据1-----*/
                    /*-----重新计算数据2-----*/
                    final CellReference cellReferenceP32 = new CellReference("P32");
                    final HSSFRow rowP32 = sheet.getRow(cellReferenceP32.getRow());
                    final HSSFCell cellP32 = rowP32.getCell(cellReferenceP32.getCol());
                    cellP32.setCellFormula(cellP32.getCellFormula());
                    /*-----重新计算数据2-----*/
                    /*-----重新计算数据3-----*/
                    final CellReference cellReferenceP37 = new CellReference("P37");
                    final HSSFRow rowP37 = sheet.getRow(cellReferenceP37.getRow());
                    final HSSFCell cellP37 = rowP37.getCell(cellReferenceP37.getCol());
                    cellP37.setCellFormula(cellP37.getCellFormula());
                    /*-----重新计算数据3-----*/
                    /*-----重新计算数据4-----*/
                    final CellReference cellReferenceP42 = new CellReference("P42");
                    final HSSFRow rowP42 = sheet.getRow(cellReferenceP42.getRow());
                    final HSSFCell cellP42 = rowP42.getCell(cellReferenceP42.getCol());
                    cellP42.setCellFormula(cellP42.getCellFormula());
                    /*-----重新计算数据4-----*/
                    /*-----重新计算数据5-----*/
                    final CellReference cellReferenceP53 = new CellReference("P53");
                    final HSSFRow rowP53 = sheet.getRow(cellReferenceP53.getRow());
                    final HSSFCell cellP53 = rowP53.getCell(cellReferenceP53.getCol());
                    cellP53.setCellFormula(cellP53.getCellFormula());
                    /*-----重新计算数据5-----*/
                    /*-----重新计算数据6-----*/
                    final CellReference cellReferenceV61 = new CellReference("V61");
                    final HSSFRow rowV61 = sheet.getRow(cellReferenceV61.getRow());
                    final HSSFCell cellV61 = rowV61.getCell(cellReferenceV61.getCol());
                    cellV61.setCellFormula(cellV61.getCellFormula());
                    /*-----重新计算数据6-----*/

                    final FileOutputStream outputStream = new FileOutputStream(excelUrl);
                    workbook.write(outputStream);
                    outputStream.flush();
                    outputStream.close();
                    statement = connection.createStatement();
                    statement.executeUpdate(sqlDeleteData);//清除数据表
                    JOptionPane.showMessageDialog(null, "写入Excel成功");
                    connection.close();
                } catch (Exception e) {
                    e.printStackTrace();
                    //JOptionPane.showMessageDialog(null, "有错误", null, JOptionPane.ERROR_MESSAGE);
                }
            }
        }
        if (excelUrl1.equals("") || mdbUrl.equals("")) {
            JOptionPane.showMessageDialog(null, "请选择文件", null, JOptionPane.WARNING_MESSAGE);
        }
    }

    //cn
    public void cn1() throws SQLException, IOException {
        String excelUrl = excelUrl1.replaceAll("\\\\", "\\\\\\\\"); //将地址转义
        POIFSFileSystem fileSystem = new POIFSFileSystem(new BufferedInputStream(new FileInputStream(excelUrl))); //使用POI读取Excel
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem); //创建HSSFWorkbook对象
        HSSFSheet sheet = workbook.getSheetAt(0); //读取第一个Sheet

        Connection connection = null; //创建Connection对象
        PreparedStatement preparedStatement = null; //创建PreparedStatement对象
        Statement statement = null;
        ResultSet resultSet = null; //创建ResultSet
        String dbUrl = "jdbc:ucanaccess://" + mdbUrl.replaceAll("\\\\", "\\\\\\\\"); //创建连接
        connection = DriverManager.getConnection(dbUrl, "username", "password"); //读取连接

        String sqlDeleteData = "delete * from TResult";

        String sqlNo400 = "select * from TResult where tableno='no400' and nindex=30";

        String sqlNo403 = "select * from TResult where tableno='no403' and nindex=26";

        String sqlNo412 = "select * from TResult where tableno='no412' and nindex=26";

        String sqlNo418 = "select * from TResult where tableno='no418' and nindex=26";

        String sqlNo424 = "select * from TResult where tableno='no424' and nindex=26";

        String sqlNo430 = "select * from TResult where tableno='no430' and nindex=28";

        String sqlNo436 = "select * from TResult where tableno='no436' and porder=9";

        /*-----读取型号单元格数据-----*/
        final Pattern pattern = Pattern.compile("[A-Z](.*?)\\d$");
        final CellReference cellReferenceP4 = new CellReference("P4");
        final HSSFRow rowP4 = sheet.getRow(cellReferenceP4.getRow());
        final HSSFCell cellP4 = rowP4.getCell(cellReferenceP4.getCol());
        String ProductType = cellP4.getStringCellValue();
        Matcher matcherProductType = pattern.matcher(ProductType.replaceAll("\\s", ""));
        /*-----读取型号单元格数据-----*/
        if (matcherProductType.find()) {
            /*-----初值-----*/
            preparedStatement = connection.prepareStatement(sqlNo400);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceA44 = new CellReference("A38");
                final HSSFRow rowA44 = sheet.getRow(cellReferenceA44.getRow());
                final HSSFCell cellA44 = rowA44.getCell(cellReferenceA44.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellA44.setCellValue(od1);
            }
            /*-----初值-----*/
            /*-----数据1-----*/
            preparedStatement = connection.prepareStatement(sqlNo403);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceG50 = new CellReference("G44");
                final HSSFRow rowG50 = sheet.getRow(cellReferenceG50.getRow());
                final HSSFCell cellG50 = rowG50.getCell(cellReferenceG50.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellG50.setCellValue(od1);
            }
            /*-----数据1-----*/
            /*-----数据2-----*/
            preparedStatement = connection.prepareStatement(sqlNo412);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceG55 = new CellReference("G49");
                final HSSFRow rowG55 = sheet.getRow(cellReferenceG55.getRow());
                final HSSFCell cellG55 = rowG55.getCell(cellReferenceG55.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellG55.setCellValue(od1);
            }
            /*-----数据3-----*/
            preparedStatement = connection.prepareStatement(sqlNo418);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceG60 = new CellReference("G54");
                final HSSFRow rowG60 = sheet.getRow(cellReferenceG60.getRow());
                final HSSFCell cellG60 = rowG60.getCell(cellReferenceG60.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellG60.setCellValue(od1);
            }
            /*-----数据4-----*/
            preparedStatement = connection.prepareStatement(sqlNo424);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceG65 = new CellReference("G59");
                final HSSFRow rowG65 = sheet.getRow(cellReferenceG65.getRow());
                final HSSFCell cellG65 = rowG65.getCell(cellReferenceG65.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellG65.setCellValue(od1);
            }
            /*-----数据5-----*/
            preparedStatement = connection.prepareStatement(sqlNo430);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceD76 = new CellReference("D70");
                final HSSFRow rowD76 = sheet.getRow(cellReferenceD76.getRow());
                final HSSFCell cellD76 = rowD76.getCell(cellReferenceD76.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellD76.setCellValue(od1);
            }
            /*-----数据5-----*/

            /*-----数据6-----*/
            preparedStatement = connection.prepareStatement(sqlNo436);
            resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                final CellReference cellReferenceD84 = new CellReference("D78");
                final HSSFRow rowD84 = sheet.getRow(cellReferenceD84.getRow());
                final HSSFCell cellD84 = rowD84.getCell(cellReferenceD84.getCol());
                Double od1 = Double.parseDouble(resultSet.getString("ODRes"));
                od1 = (double) Math.round(od1 * 10000) / 10000;
                cellD84.setCellValue(od1);
            }
            /*-----重新计算初值1-----*/
            final CellReference cellReferenceP44 = new CellReference("P38");
            final HSSFRow rowP44 = sheet.getRow(cellReferenceP44.getRow());
            final HSSFCell cellP44 = rowP44.getCell(cellReferenceP44.getCol());
            cellP44.setCellFormula(cellP44.getCellFormula());
            /*-----重新计算初值1-----*/
            /*-----重新计算数据1-----*/
            final CellReference cellReferenceP50 = new CellReference("P44");
            final HSSFRow rowP50 = sheet.getRow(cellReferenceP50.getRow());
            final HSSFCell cellP50 = rowP50.getCell(cellReferenceP50.getCol());
            cellP50.setCellFormula(cellP50.getCellFormula());
            /*-----重新计算数据1-----*/
            /*-----重新计算数据2-----*/
            final CellReference cellReferenceP55 = new CellReference("P49");
            final HSSFRow rowP55 = sheet.getRow(cellReferenceP55.getRow());
            final HSSFCell cellP55 = rowP55.getCell(cellReferenceP55.getCol());
            cellP55.setCellFormula(cellP55.getCellFormula());
            /*-----重新计算数据2-----*/
            /*-----重新计算数据3-----*/
            final CellReference cellReferenceP60 = new CellReference("P54");
            final HSSFRow rowP60 = sheet.getRow(cellReferenceP60.getRow());
            final HSSFCell cellP60 = rowP60.getCell(cellReferenceP60.getCol());
            cellP60.setCellFormula(cellP60.getCellFormula());
            /*-----重新计算数据3-----*/
            /*-----重新计算数据4-----*/
            final CellReference cellReferenceP65 = new CellReference("P59");
            final HSSFRow rowP65 = sheet.getRow(cellReferenceP65.getRow());
            final HSSFCell cellP65 = rowP65.getCell(cellReferenceP65.getCol());
            cellP65.setCellFormula(cellP65.getCellFormula());
            /*-----重新计算数据4-----*/
            /*-----重新计算数据5-----*/
            final CellReference cellReferenceP76 = new CellReference("P70");
            final HSSFRow rowP76 = sheet.getRow(cellReferenceP76.getRow());
            final HSSFCell cellP76 = rowP76.getCell(cellReferenceP76.getCol());
            cellP76.setCellFormula(cellP76.getCellFormula());
            final CellReference cellReferenceR76 = new CellReference("R70");
            final HSSFRow rowR76 = sheet.getRow(cellReferenceR76.getRow());
            final HSSFCell cellR76 = rowR76.getCell(cellReferenceR76.getCol());
            cellR76.setCellFormula(cellR76.getCellFormula());
            /*-----重新计算数据5-----*/
            /*-----重新计算数据6-----*/
            final CellReference cellReferenceV84 = new CellReference("V78");
            final HSSFRow rowV84 = sheet.getRow(cellReferenceV84.getRow());
            final HSSFCell cellV84 = rowV84.getCell(cellReferenceV84.getCol());
            cellV84.setCellFormula(cellV84.getCellFormula());
            /*-----重新计算数据6-----*/
        }

        final FileOutputStream outputStream = new FileOutputStream(excelUrl);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        statement = connection.createStatement();
        statement.executeUpdate(sqlDeleteData);//清除数据表
        JOptionPane.showMessageDialog(null, "写入Excel成功");
        connection.close();
    }

}

class ButtonListener implements ActionListener {
    private Report report;

    public ButtonListener(Report report) {
        super();
        this.report = report;
    }

    public void actionPerformed(ActionEvent e) {
        if (e.getActionCommand().equals("获取Excel报告位置")) {
            report.showFileOpenDialog();
        }
        if (e.getActionCommand().equals("获取数据库Data位置")) {
            report.showDBOpenDialog();
        }
        if (e.getActionCommand().equals("填写报告")) {
            try {
                report.showExcel();
            } catch (SQLException ex) {
                ex.printStackTrace();
            }
        }
    }
}
