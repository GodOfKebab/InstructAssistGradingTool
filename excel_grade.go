package main

import (
	"errors"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
)

func main() {
	var excelPath = flag.String("f", "ProjectGrades.xlsx", "Filepath for the grades excel file")
	flag.Parse()

	f, errFile := excelize.OpenFile(*excelPath)
	if errFile != nil {
		log.Fatal(errFile)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// Get 2D-array 'sheet' of type string
	sheet, errSheet := f.GetRows("Sheet1")
	if errSheet != nil {
		log.Fatal(errSheet)
	}

	// In case comments cell is empty, default to these
	defaults, defaultsErr := populateDefaultComments(sheet)
	if defaultsErr != nil {
		log.Fatal(defaultsErr)
	}

	// Delete folder 'attachments'
	if err := os.RemoveAll("attachments"); os.IsNotExist(err) {
		fmt.Println("Can't delete folder 'attachments'!")
		log.Fatal(err)
	}

	// Create folder 'attachments'
	if err := os.Mkdir("attachments", os.ModePerm); err != nil {
		fmt.Println("Can't create folder 'attachments' to save comments files to!")
		log.Fatal(err)
	}

	// Get rubric field type: Task/Q
	rubricFieldType, errRubricField := getRubricFieldType(sheet)
	if errRubricField != nil {
		log.Fatal(errRubricField)
	}

	// Create the .txt files
	for _, row := range sheet[1:] {
		csfError := createStudentAttachment(row, defaults, rubricFieldType, "attachments")
		if csfError != nil {
			fmt.Println(csfError)
		}

	}
}

func populateDefaultComments(sheet [][]string) ([]string, error) {
	var defaultCommentsCells []string
	defaultCommentsStartPos := [...]int{0, 0}
	for i, row := range sheet {
		for j, cell := range row {
			if cell == "Default Comments:" {
				defaultCommentsStartPos[0] = i + 1
				defaultCommentsStartPos[1] = j + 1
			}
		}
	}
	if defaultCommentsStartPos[0] == 0 && defaultCommentsStartPos[1] == 0 {
		return make([]string, 0), errors.New("couldn't find cell: Default Comments")
	}

	for _, row := range sheet[defaultCommentsStartPos[0]:] {
		if len(row) > defaultCommentsStartPos[1] {
			defaultCommentsCells = append(defaultCommentsCells, row[defaultCommentsStartPos[1]])
		}
	}
	return defaultCommentsCells, nil
}

func getRubricFieldType(sheet [][]string) (string, error) {
	for i, row := range sheet {
		for j, cell := range row {
			if cell == "Rubric Field Type:" {
				return sheet[i+1][j], nil
			}
		}
	}
	return "", errors.New("couldn't find cell: Rubric Field Type")
}

func createStudentAttachment(row []string, defaults []string, rubricFieldType string, folderName string) error {
	content := []byte("Comments:\n\n")

	// Per field comments
	for i, defaultComment := range defaults {
		var comment string
		if row[3+i*2] != "" {
			comment = row[3+i*2]
		} else {
			comment = defaultComment
		}
		commentLine := fmt.Sprintf("%s%d: %s\n\n", rubricFieldType, i+1, comment)
		content = append(content, []byte(commentLine)...)
	}

	// Overall comments
	if len(row) > 5+(len(defaults)-1)*2 {
		if row[5+(len(defaults)-1)*2] != "" {
			overallCommentLine := fmt.Sprintf("%s\n\n", row[5+(len(defaults)-1)*2])
			content = append(content, []byte(overallCommentLine)...)
		}
	}

	// Per field scores
	content = append(content, []byte("Scores:\n")...)
	for i, _ := range defaults {
		commentLine := fmt.Sprintf("\t%s%d: %s", rubricFieldType, i+1, row[2+i*2])
		content = append(content, []byte(commentLine)...)
	}

	// Total Score
	commentLine := fmt.Sprintf("\n\nTotal Score: %s", row[4+(len(defaults)-1)*2])
	content = append(content, []byte(commentLine)...)

	attachmentFilePath := fmt.Sprintf("%s/%s.txt", folderName, row[1])
	return os.WriteFile(attachmentFilePath, content, 0644)
}
