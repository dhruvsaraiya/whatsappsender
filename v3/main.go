package main

import (
    "encoding/gob"
    "fmt"
    // "github.com/Baozisoftware/qrcode-terminal-go"
    "github.com/Rhymen/go-whatsapp"
    "github.com/tealeg/xlsx"
    "os"
    "os/user"
    "time"
    "strings"
    "unicode"
    "strconv"
    //"bufio"
)

type NUMBER map[string]interface{}

func main() {
    go_o, _ :=  os.OpenFile("go_out.txt", os.O_CREATE|os.O_RDWR, 0666)
    os.Stdout = go_o
    go_e, _ :=  os.OpenFile("go_err.txt", os.O_CREATE|os.O_RDWR, 0666)
    os.Stderr = go_e
    
    //create new WhatsApp connection
    if (len(os.Args) != 3) && (len(os.Args) != 5){
        fmt.Fprintf(os.Stderr, "invalid arguments.....")
        return
    }
    excel_name := os.Args[1]
    var image_flag bool = false
    var image_name string
    var option string
    var start_time string

    if (len(os.Args) == 5){
        image_flag = true
        option = os.Args[3]
        start_time = os.Args[4]
    } else {
        start_time = os.Args[2]
    }
    if image_flag{
        image_name = os.Args[2]
    }

    _, err_e := os.Open(excel_name)
    if err_e != nil{
        fmt.Fprintf(os.Stderr, "error opening excel : %v\n", err_e)
        return
    }
    if image_flag{
        _, err_i := os.Open(image_name)
        if err_i != nil{
            fmt.Fprintf(os.Stderr, "error opening image : %v\n", err_i)
            return
        }
    }

    wac, err := whatsapp.NewConn(15 * time.Second)
    
    if err != nil {
        fmt.Fprintf(os.Stderr, "error creating connection: %v\n", err)
        return
    }

    err = login(wac)
    if err != nil {
        fmt.Fprintf(os.Stderr, "error logging in: %v\n", err)
        return
    }

    <-time.After(3 * time.Second)

    fmt.Printf("START TIME = %v \n", time.Now().Format("2006-01-02 15:04:05"))
    // start_time := time.Now().Format("2006-01-02 15:04:05")
    start_time = strings.Replace(start_time, " ", "_", -1)
    start_time = strings.Replace(start_time, "-", "_", -1)
    start_time = strings.Replace(start_time, ":", "_", -1)
    success_file := strings.Join([]string{"success", start_time, ".xlsx"}, "")
    error_file := strings.Join([]string{"error", start_time, ".xlsx"}, "")
    myself, error := user.Current()
    if error != nil {
        panic(error)
    }
    homedir := myself.HomeDir
    desktop := homedir + "/Desktop/" 
    success_file = desktop + success_file
    error_file = desktop + error_file
    
    //fmt.Printf("value of success_file = %v ", success_file)

    numbers := getMessages(excel_name, error_file)

    fmt.Printf("EXCEL READ AT TIME = %v \n", time.Now().Format("2006-01-02 15:04:05"))
    var s_file *xlsx.File
    var s_sheet *xlsx.Sheet
    //var e_row *xlsx.Row
    var s_cell *xlsx.Cell
    var e_file *xlsx.File
    var e_sheet *xlsx.Sheet
    //var e_row *xlsx.Row
    var e_cell *xlsx.Cell
    var error_columns []string
    error_columns = append(error_columns, "phone", "message", "reason")
    //var e_row_count int = 1
    var success_columns []string
    success_columns = append(success_columns, "phone", "message")
    s_file, s_err := xlsx.OpenFile(success_file)
    
    if s_err != nil {
        fmt.Printf("success file not found.... %v\n", s_err)
        s_file = xlsx.NewFile()
        s_sheet, _ = s_file.AddSheet("sent")
        col_name_row := s_sheet.AddRow()
        for _, column := range success_columns{
            s_cell = col_name_row.AddCell()
            s_cell.Value = column
        }
    } else {
        s_sheet = s_file.Sheets[0]
    }

    temp_count := 0
    sf, err :=  os.OpenFile("status.txt", os.O_CREATE|os.O_RDWR, 0666)
    if err == nil{
        sf.WriteString("0")
    }
    sf.Close();

    e_file, err_f := xlsx.OpenFile(error_file)
    if err_f != nil {
        fmt.Printf("ERROR %v\n", err)
        fmt.Printf("error file not found.... %v\n", s_err)
        e_file = xlsx.NewFile()
        e_sheet, _ = e_file.AddSheet("failed")
        e_col_name_row := e_sheet.AddRow()
        for _, column := range error_columns{
            e_cell = e_col_name_row.AddCell()
            e_cell.Value = column
        }
    } else {
    	e_sheet = e_file.Sheets[0]
    }

    for count, n := range numbers {
        //fmt.Printf("value of x = %v at %d\n", n,i)
        // fmt.Printf("name = %v ", n["name"])
        fmt.Printf("phone = %v \n", n["phone"])
        //fmt.Printf("value of message = %v \n", n["message"])
        
        var at string = "@s.whatsapp.net"
        var id string = strings.Join([]string{n["phone"].(string), at}, "")
        
        //fmt.Printf("value of string = %v \n", id)

        if image_flag{
            img, err := os.Open(image_name)
            if err != nil {
                //fmt.Fprintf(os.Stderr, "error reading file: %v\n", err)
                // fmt.Printf("ERROR for user : %v, in opening image\n", n["name"].(string))
                fmt.Printf("ERROR in opening image for %v\n", n["phone"])
                //os.Exit(1)
            }
            
            if option == "1"{
                msg_t := whatsapp.TextMessage{
                    Info: whatsapp.MessageInfo{
                        RemoteJid: id,
                    },
                    Text: n["message"].(string),
                }
                err = wac.Send(msg_t)

                fmt.Printf("ERROR for %v\n", err)

                msg_i := whatsapp.ImageMessage{
                    Info: whatsapp.MessageInfo{
                    RemoteJid: id,
                    },
                    Type: "image/jpeg",
                    Content: img,
                }
                err = wac.Send(msg_i)
                
                fmt.Printf("ERROR for %v\n", err)

            } else if option == "2"{
                msg_i := whatsapp.ImageMessage{
                    Info: whatsapp.MessageInfo{
                    RemoteJid: id,
                    },
                    Type: "image/jpeg",
                    Content: img,
                }
                err = wac.Send(msg_i)
                fmt.Printf("ERROR for %v\n", err)

                msg_t := whatsapp.TextMessage{
                    Info: whatsapp.MessageInfo{
                        RemoteJid: id,
                    },
                    Text: n["message"].(string),
                }
                err = wac.Send(msg_t)
                fmt.Printf("ERROR for %v\n", err)
            } else if option == "3"{

                msg_c := whatsapp.ImageMessage{
                    Info: whatsapp.MessageInfo{
                        RemoteJid: id,
                    },
                    Type: "image/jpeg",
                    Caption: n["message"].(string),
                    Content: img,
                }
                err = wac.Send(msg_c)
                fmt.Printf("ERROR for %v\n", err)
            }
        } else {
            msg := whatsapp.TextMessage{
                Info: whatsapp.MessageInfo{
                    RemoteJid: id,
                },
                Text: n["message"].(string),
            }
            err = wac.Send(msg)
            fmt.Printf("ERROR for %v\n", err)
        }

        error_flag := 0
        var error_string string
        if err != nil{
        	error_string = err.Error()
        	if strings.Contains(error_string, "sending message timed out"){
        		error_flag = 0
        	} else {
        		error_flag = 1
        	}

        }
        if error_flag == 1{
	        new_error_row := e_sheet.AddRow()
	        for _, column := range error_columns{
	            e_cell = new_error_row.AddCell()
	            if column != "reason"{
		            e_cell.Value = n[column].(string)
	            } else {
	            	e_cell.Value = error_string
	            }
	        }
	        e_file.Save(error_file)
        } else {
	        new_success_row := s_sheet.AddRow()
	        for _, column := range success_columns{
	            s_cell = new_success_row.AddCell()
	            s_cell.Value = n[column].(string)
	        }
	        s_file.Save(success_file)        	
        }

        temp_count ++;
        var per int;
        if temp_count >= 5{
            sf, err :=  os.OpenFile("status.txt", os.O_CREATE|os.O_RDWR, 0666)
            if err == nil{
                per = (int)(float64(count+1)/float64(len(numbers))*100)
                per_str := strconv.Itoa(per)
                fmt.Printf("Status ==>> %v,%v,%v,%s", count,len(numbers),per,per_str)
                sf.WriteString(per_str)
            }
            sf.Close();
            temp_count = 0
        }
        // if err != nil {
        //     // fmt.Fprintf(os.Stderr, "error sending file: %v\n", err)
        //     new_success_row := s_sheet.AddRow()
        //     for _, column := range success_columns{
        //         s_cell = new_success_row.AddCell()
        //         s_cell.Value = n[column].(string)
        //     }
        //     s_file.Save(success_file)
        //     // fmt.Printf("ERROR for user : %v\n", n["phone"].(string))
        //     // fmt.Printf("this does not always mean message is not sent, please check out manually, it will be in SUCCESS LOG though.\n")
        //     //os.Exit(1)
        // } else {
        //     new_success_row := s_sheet.AddRow()
        //     for _, column := range success_columns{
        //         s_cell = new_success_row.AddCell()
        //         s_cell.Value = n[column].(string)
        //     }
        //     s_file.Save(success_file)
        // }
    }
    fmt.Printf("END TIME = %v \n", time.Now().Format("2006-01-02 15:04:05"))
}

func login(wac *whatsapp.Conn) error {
	
    my_e, _ :=  os.OpenFile("my_err.txt", os.O_CREATE|os.O_RDWR, 0666)
    //load saved session
    session, err := readSession()
    if err == nil {
        //restore session
        session, err = wac.RestoreSession(session)
        if err != nil {
            my_e.WriteString("restore")
            my_e.Close()
            return fmt.Errorf("restoring failed: %v\n", err)
        }
    } else {
        //no saved session -> regular login
        qr := make(chan string)

        // go func() {
        //     terminal := qrcodeTerminal.New()
        //     terminal.Get(<-qr).Print()
        // }()

        go func() {
            qrstring := <-qr
            f, _ :=  os.OpenFile("qrcode.txt", os.O_CREATE|os.O_RDWR, 0666)
            f.WriteString(qrstring)
            f.Close()
            fmt.Printf("QR string ==>> %v", qrstring)
        }()

        session, err = wac.Login(qr)

        if err != nil {
            my_e.WriteString("qrcode")
            my_e.Close()
            return fmt.Errorf("error during login: %v\n", err)
        }
    }

    //save session
    err = writeSession(session)
    if err != nil {
        return fmt.Errorf("error saving session: %v\n", err)
    }
    return nil
}

func readSession() (whatsapp.Session, error) {
    session := whatsapp.Session{}
    file, err := os.Open(os.TempDir() + "/whatsappSession.gob")
    if err != nil {
        return session, err
    }
    defer file.Close()
    decoder := gob.NewDecoder(file)
    err = decoder.Decode(&session)
    if err != nil {
        return session, err
    }
    return session, nil
}

func writeSession(session whatsapp.Session) error {
    file, err := os.Create(os.TempDir() + "/whatsappSession.gob")
    if err != nil {
        return err
    }
    defer file.Close()
    encoder := gob.NewEncoder(file)
    err = encoder.Encode(session)
    if err != nil {
        return err
    }
    return nil
}

/*
func getMessages() []NUMBER{
    
    var numbers []string
    numbers = append(numbers, "919969050933")
    numbers = append(numbers, "919167470511")
    numbers = append(numbers, "919910844456")
    n1 := NUMBER{"name": "yash", "phone": "919167470511", "message": "hello yash from api"}
    n2 := NUMBER{"name": "rahul", "phone": "919969050933", "message": "hello rahul from api"}
    n3 := NUMBER{"name": "sachin", "phone": "919910844456", "message": "hello sachin from api"}
    n4 := NUMBER{"name": "abhinav", "phone": "919887360137", "message": "hello abhinav from api"}
    numbers = append(numbers, n1, n2, n3, n4)

    excelData := readExcel("whatsapp.xlsx")
    //fmt.Printf("EXCEL %v\n", excelData)
    return excelData

}
*/

func getMessages(fnm string, error_file string) []NUMBER{
    var data []NUMBER
    var e_file *xlsx.File
    var e_sheet *xlsx.Sheet
    //var e_row *xlsx.Row
    var e_cell *xlsx.Cell
    //var e_row_count int = 1
    var error_columns []string
    error_columns = append(error_columns, "phone", "message", "reason")
    e_file, err := xlsx.OpenFile(error_file)
    
    if err != nil {
        fmt.Printf("error file not found.... %v\n", err)
        e_file = xlsx.NewFile()
        e_sheet, _ = e_file.AddSheet("failed")
        col_name_row := e_sheet.AddRow()
        for _, column := range error_columns{
            e_cell = col_name_row.AddCell()
            e_cell.Value = column
        }
    } else {
        e_sheet = e_file.Sheets[0]
    }

    //e_row_count = len(e_sheet.Rows) + 1


    xlFile, err := xlsx.OpenFile(fnm)
    if err != nil {
        fmt.Printf("ERROR %v\n", err)
        return nil
    }
    
    var columns []string
    first_row := xlFile.Sheets[0].Rows[0]
    fmt.Printf("ROW : %v\n", first_row)
    
    for _, cell := range first_row.Cells{
        columns = append(columns, cell.String())
    }
    
    fmt.Printf("COLUMNS : %v\n", columns)
    var skip_flag bool
    for _, row := range xlFile.Sheets[0].Rows[1:]{
        n1 := NUMBER{}
        skip_flag = false
        for col_index, column := range columns{
            n1[column] = row.Cells[col_index].String()
            if column == "phone"{
                if len(n1[column].(string)) != 10 {
                    skip_flag = true
                    //break
                }
            }
            if column == "phone"{
                if !isInt(n1[column].(string)){
                    skip_flag = true
                    //break
                }
            }
            if column == "phone"{
                n1[column] = strings.Join([]string{"91", n1["phone"].(string)}, "")
            }
        }
        if skip_flag{
            fmt.Printf("error..........\n")
            //write to error excel file
            new_error_row := e_sheet.AddRow()
            for _, column := range error_columns{
                e_cell = new_error_row.AddCell()
                e_cell.Value = n1[column].(string)
            }
            e_file.Save(error_file)
        } else {
            data = append(data, n1)
        }
    }
    return data
}

func isInt(s string) bool {
    for _, c := range s {
        if !unicode.IsDigit(c) {
            return false
        }
    }
    return true
}
