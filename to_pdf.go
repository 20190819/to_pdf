package main

import (
	"fmt"
	"os"
	"os/exec"
	"path"
	"strings"
)

/**
*@tips libreoffice 转换指令：
* 安装 libreoffice 参考 https://www.cnblogs.com/qlqwjy/p/9846904.html
* libreoffice6.2 invisible --convert-to pdf csDoc.doc --outdir /home/[转出目录]
* soffice invisible --convert-to pdf csDoc.doc --outdir /home/[转出目录]	// windows
*
* @function 实现文档类型转换为pdf或html
* @param
*	  command:libreofficed的命令(具体以版本为准)；win：soffice； linux：libreoffice6.2
*     fileSrcPath:转换文件的路径
*     fileOutDir:转换后文件存储目录
*     converterType：转换的类型 pdf 或 html
* @return fileOutPath 转换成功生成的文件的路径 error 转换错误
 */
func FuncDocs2Pdf(command string, fileSrcPath string, fileOutDir string, converterType string) (fileOutPath string, error error) {
	//校验fileSrcPath
	srcFile, erByOpenSrcFile := os.Open(fileSrcPath)
	if erByOpenSrcFile != nil && os.IsNotExist(erByOpenSrcFile) {
		return "", erByOpenSrcFile
	}
	//如文件输出目录fileOutDir不存在则自动创建
	outFileDir, erByOpenFileOutDir := os.Open(fileOutDir)
	if erByOpenFileOutDir != nil && os.IsNotExist(erByOpenFileOutDir) {
		erByCreateFileOutDir := os.MkdirAll(fileOutDir, os.ModePerm)
		if erByCreateFileOutDir != nil {
			fmt.Println("File ouput dir create error.....", erByCreateFileOutDir.Error())
			return "", erByCreateFileOutDir
		}
	}
	//关闭流
	defer func() {
		_ = srcFile.Close()
		_ = outFileDir.Close()
	}()
	//convert
	cmd := exec.Command(command, "--invisible", "--convert-to", converterType,
		fileSrcPath, "--outdir", fileOutDir)
	byteByStat, errByCmdStart := cmd.Output()
	//命令调用转换失败
	if errByCmdStart != nil {
		return "", errByCmdStart
	}
	//success
	fileOutPath = fileOutDir + "/" + strings.Split(path.Base(fileSrcPath), ".")[0]
	if converterType == "html" {
		fileOutPath += ".html"
	} else {
		fileOutPath += ".pdf"
	}
	fmt.Println("文件转换成功...", string(byteByStat))
	return fileOutPath, nil
}

func main() {
	var command string = "soffice"
	// var fileSrcPath string = "./liangyang.doc"
	var fileSrcPath string = "./test.docx"
	var fileOutDir string = "./pdf_files"
	var converterType string = "pdf"

	pdf_path, err := FuncDocs2Pdf(command, fileSrcPath, fileOutDir, converterType)

	if err == nil {
		fmt.Println("pdf_path", pdf_path)
	} else {
		fmt.Println("has error", err)
	}
}
