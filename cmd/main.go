package main

import (
	"fmt"
	"log"

	"github.com/nguyenthenguyen/docx"
)

func main() {
	// Открываем шаблон Word-документа
	r, err := docx.ReadDocxFile("./source/act.docx")
	if err != nil {
		log.Fatalf("Ошибка открытия документа: %v", err)
	}
	defer r.Close()

	// Получаем содержимое документа
	doc := r.Editable()

	// Переменные для замены
	variables := map[string]string{
		"DateDocument": "31.10.2022",
		//"DescriptionFirstPosition": "ЕРП системе «Интеграция с банком Тинькофф Рассрочки»",
		//	"DescriptionSecondPosition": "ЕРП системе «Интеграция с поатёжной системой Сбер Эквайринг",
		//	"DescriptionThirdPosition":  "ЕРП системе «Интеграция с платёжной системой PayKeeper»",
		//	"PriceFirstPosition":        "50 000",
		//	"PriceSecondPosition":       "63 000",
		////	"PriceThirdPosition":        "50 000»",
		//	"DateDocument":              "31.10.2022",
		//	"TotalPrice":                "163 000,00",
		//"TotalPriceText":            "сто шестьдесят три тысячи 00 копеек",
	}

	// Заменяем переменные в тексте
	for key, value := range variables {
		placeholder := fmt.Sprintf("{%s}", key)

		err := doc.Replace(placeholder, value, -1)
		if err != nil {
			fmt.Println(err)
			return
		}

		fmt.Printf("Проверка текста: %s\n", key)
		fmt.Printf("Заменяем %s на %s\n", placeholder, value)
	}

	// Сохраняем измененный документ
	outputFile := "./dist/act.docx"
	err = doc.WriteToFile(outputFile)
	if err != nil {
		log.Fatalf("Ошибка сохранения документа: %v", err)
	}

	fmt.Printf("Документ успешно сохранен: %s\n", outputFile)
}
