package main

import (
	"fmt"
	"github.com/nguyenthenguyen/docx"
	"log"
	"os"
	"time"
)

func main() {

	reasonList := map[string]map[int]map[string]string{
		"30.06.2022": {
			0: {
				"price":       "100 000",
				"description": "ЕРП системе «Интеграция с банком Тинькофф Рассрочки»",
			},
			1: {
				"price":       "60 883",
				"description": "ЕРП системе «Интеграция с платёжной системой Сбер Эквайринг",
			},
			2: {
				"price":       "80 000",
				"description": "ЕРП системе «Интеграция с платёжной системой PayKeeper»",
			},
			3: {
				"total_price":      "240 883,00",
				"total_price_text": "двести сорок тысяч восемьсот восемьдесят три рублей 00 копеек.",
			},
		},
		"31.07.2022": {
			0: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Реестр Возвратов»",
			},
			1: {
				"price":       "34 190",
				"description": "ЕРП системе Разработка модуля «Поиск информации по ученику»",
			},
			2: {
				"price":       "40 000",
				"description": "ЕРП системе Разработка модуля «Отзывы об уроках»",
			},
			3: {
				"total_price":      "134 190,00",
				"total_price_text": "сто тридцать четыре тысячи сто девяносто рубля 00 копеек.",
			},
		},
		"31.10.2022": {
			0: {
				"price":       "30 000",
				"description": "ЕРП системе Разработка модуля «Роялти»",
			},
			1: {
				"price":       "70 000",
				"description": "ЕРП системе Разработка модуля «Сотрудники»",
			},
			2: {
				"price":       "63 000",
				"description": "ЕРП системе Разработка модуля «Школы»",
			},
			3: {
				"total_price":      "163 000,00",
				"total_price_text": "сто шестьдесят три тысячи рублей 00 копеек.",
			},
		},
		"30.11.2022": {
			0: {
				"price":       "100 000",
				"description": "ЕРП системе Разработка модуля «Курсы»",
			},
			1: {
				"price":       "60 884",
				"description": "ЕРП системе Разработка модуля «Уроки»",
			},
			2: {
				"price":       "80 000",
				"description": "ЕРП системе Разработка модуля «Программы»",
			},
			3: {
				"total_price":      "240 884,00",
				"total_price_text": "двести сорок тысяч восемьсот восемьдесят четыре рубля 00 копеек.",
			},
		},
		"30.12.2022": {
			0: {
				"price":       "40 000",
				"description": "ЕРП системе Разработка модуля «Группы»",
			},
			1: {
				"price":       "35 955",
				"description": "ЕРП системе Разработка модуля «Уроки»",
			},
			2: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Расписание»",
			},
			3: {
				"total_price":      "125 955,00",
				"total_price_text": "сто двадцать пять тысяч девятьсот пятьдесят пять рублей 00 копеек.",
			},
		},
		"31.03.2023": {
			0: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Расписание»",
			},
			1: {
				"price":       "63 000",
				"description": "ЕРП системе Разработка модуля «Задачи»",
			},
			2: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «История звонков»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"30.04.2023": {
			0: {
				"price":       "43 000",
				"description": "ЕРП системе Разработка модуля «История звонков»",
			},
			1: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Вконтакте»",
			},
			2: {
				"price":       "70 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Facebook»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.05.2023": {
			0: {
				"price":       "100 000",
				"description": "ЕРП системе Разработка модуля «Интеграция SMS центр»",
			},
			1: {
				"price":       "60 883",
				"description": "ЕРП системе Разработка модуля «Интеграция SMS центр»",
			},
			2: {
				"price":       "80 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Телефония Sipuni»",
			},
			3: {
				"total_price":      "254 228,00",
				"total_price_text": "двести пятьдесят четыре тысячи двести двадцать восемь рублей 00 копеек.",
			},
		},
		"30.06.2023": {
			0: {
				"price":       "73 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Zoom»",
			},
			1: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Jitsi»",
			},
			2: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Интеграция Google Meet»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.07.2023": {
			0: {
				"price":       "56 000",
				"description": "ЕРП системе Разработка модуля «Курсы»",
			},
			1: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Архив лидов»",
			},
			2: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Отчёт по выручке партнёров»",
			},
			3: {
				"total_price":      "176 000,00",
				"total_price_text": "сто семьдесят шесть тысяч рублей 00 копеек.",
			},
		},
		"31.08.2023": {
			0: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Отчёт по Колл-Центру»",
			},
			1: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Отчёт по открытым урокам»",
			},
			2: {
				"price":       "63 000",
				"description": "ЕРП системе Разработка модуля «Отчёт по активным школам»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"30.09.2023": {
			0: {
				"price":       "73 000",
				"description": "ЕРП системе Разработка модуля «Отчёт конверсия по КЦ»",
			},
			1: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Отчёт конверсия по воронке продаж»",
			},
			2: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Общий отчёт по партнёру»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.10.2023": {
			0: {
				"price":       "100 000",
				"description": "ЕРП системе Разработка модуля «Ученики»",
			},
			1: {
				"price":       "65 512",
				"description": "ЕРП системе Разработка модуля «Лиды»",
			},
			2: {
				"price":       "100 000",
				"description": "ЕРП системе Разработка модуля «Архив лидов»",
			},
			3: {
				"total_price":      "265 512,00",
				"total_price_text": "двести шестьдесят пять тысяч пятьсот двенадцать рублей 00 копеек.",
			},
		},
		"30.11.2023": {
			0: {
				"price":       "80 000",
				"description": "ЕРП системе Разработка модуля «Выплата франчайзи»",
			},
			1: {
				"price":       "23 000",
				"description": "ЕРП системе Разработка модуля «Банковские платёжки»",
			},
			2: {
				"price":       "80 000",
				"description": "ЕРП системе Разработка модуля «Выгрузка чеков»",
			},
			3: {
				"total_price":      "183 000,00",
				"total_price_text": "сто восемьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.12.2023": {
			0: {
				"price":       "80 000",
				"description": "ЕРП системе Разработка модуля «Кабинет методиста»",
			},
			1: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Кабинет ученика»",
			},
			2: {
				"price":       "33 000",
				"description": "ЕРП системе Разработка модуля «Кабинет партнёра»",
			},
			3: {
				"total_price":      "163 000,00",
				"total_price_text": "сто шестьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.01.2024": {
			0: {
				"price":       "70 000",
				"description": "ЕРП системе Разработка модуля «Кабинет менеджера»",
			},
			1: {
				"price":       "51 000",
				"description": "ЕРП системе Разработка модуля «Кабинет преподавателя»",
			},
			2: {
				"price":       "52 000",
				"description": "ЕРП системе Разработка модуля «Кабинет администратора»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.03.2024": {
			0: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Самостоятельная работа»",
			},
			1: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Презентация для учёбы»",
			},
			2: {
				"price":       "63 000",
				"description": "ЕРП системе Разработка модуля «Лиды/Воронки продаж»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"30.04.2024": {
			0: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Ценообразование»",
			},
			1: {
				"price":       "90 600",
				"description": "ЕРП системе Разработка модуля «Пакеты занятий»",
			},
			2: {
				"price":       "50 000",
				"description": "ЕРП системе Разработка модуля «Новостная лента»",
			},
			3: {
				"total_price":      "190 600,00",
				"total_price_text": "сто девяносто тысяч шестьсот рублей 00 копеек.",
			},
		},
		"31.05.2024": {
			0: {
				"price":       "83 000",
				"description": "ЕРП системе Разработка модуля «Статусы учеников»",
			},
			1: {
				"price":       "125 871",
				"description": "ЕРП системе Разработка модуля «Журнал посещений»",
			},
			2: {
				"price":       "60 000",
				"description": "ЕРП системе Разработка модуля «Отработки уроков»",
			},
			3: {
				"total_price":      "268 871,00",
				"total_price_text": "двести шестьдесят восемь тысяч восемьсот семьдесят один рубль 00 копеек.",
			},
		},
		"30.06.2024": {
			0: {
				"price":       "53 000",
				"description": "ЕРП системе Разработка модуля «Перенос занятий»»",
			},
			1: {
				"price":       "100 000",
				"description": "ЕРП системе Разработка модуля «Выставление счетов»",
			},
			2: {
				"price":       "20 000",
				"description": "ЕРП системе Разработка модуля «Возврат средств за уроки»",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.07.2024": {
			0: {
				"price":       "50 000",
				"description": "ЕРП системе «Интеграция с ЮКасса",
			},
			1: {
				"price":       "80 000",
				"description": "ЕРП система «Интеграция с Атол»",
			},
			2: {
				"price":       "43 000",
				"description": "Доработка раздела отчеты",
			},
			3: {
				"total_price":      "173 000,00",
				"total_price_text": "сто семьдесят три тысячи рублей 00 копеек.",
			},
		},
		"31.08.2024": {
			0: {
				"price":       "20 000",
				"description": "Изменение функционала начисления КодКоинов",
			},
			1: {
				"price":       "39 970",
				"description": "Доработка проблем ЕРП системы",
			},
			2: {
				"price":       "115 000",
				"description": "Интеграция скоринга клиентов",
			},
			3: {
				"total_price":      "174 970,00",
				"total_price_text": "сто семьдесят четыре тысячи девятьсот семьдесят рублей 00 копеек.",
			},
		},
		"30.09.2024": {
			0: {
				"price":       "38 786",
				"description": "Интеграция с AMO CRM",
			},
			1: {
				"price":       "85 690",
				"description": "ЕРП система «Доработка раздела выплата партнёрам»",
			},
			2: {
				"price":       "65 000",
				"description": "ЕРП система «Работы по безопасности серверов ПО»",
			},
			3: {
				"total_price":      "189 476,00",
				"total_price_text": "сто восемьдесят девять тысяч четыреста семьдесят шесть рублей 00 копеек.",
			},
		},
		"31.12.2024": {
			0: {
				"price":       "104 000",
				"description": "Интеграция с Tinkoff рассрочки (выгрузка по API отчётов)",
			},
			1: {
				"price":       "50 232",
				"description": "ЕРП система «Доработка раздела Отчеты»",
			},
			2: {
				"price":       "50 000",
				"description": "ЕРП система «Функционал назначения КодКоинов для преподавателей»",
			},
			3: {
				"total_price":      "204 232,00",
				"total_price_text": "двести четыре тысячи двести тридцать два рубля 00 копеек.",
			},
		},
		"31.01.2025": {
			0: {
				"price":       "75 320",
				"description": "Разработка Telegram бота",
			},
			1: {
				"price":       "75 000",
				"description": "Разработка функционала по формированию оферт для пользователей",
			},
			2: {
				"price":       "55 000",
				"description": "Разработка ЕРП системы, для новой школы компании",
			},
			3: {
				"total_price":      "205 320,00",
				"total_price_text": "двести пять тысяч триста двадцать рублей 00 копеек.",
			},
		},
	}

	for i, arr := range reasonList {
		writeDoc(i, arr)
	}

}

func writeDoc(dateText string, arr map[int]map[string]string) {
	// Открываем шаблон Word-документа
	rAct, err := docx.ReadDocxFile("./source/act.docx")
	if err != nil {
		log.Fatalf("Ошибка открытия документа: %v", err)
	}

	rInvoice, err := docx.ReadDocxFile("./source/invoice.docx")
	if err != nil {
		log.Fatalf("Ошибка открытия документа: %v", err)
	}

	rOrder, err := docx.ReadDocxFile("./source/order.docx")
	if err != nil {
		log.Fatalf("Ошибка открытия документа: %v", err)
	}

	defer rAct.Close()
	defer rInvoice.Close()
	defer rOrder.Close()

	// Получаем содержимое документа
	docAct := rAct.Editable()
	docInvoice := rInvoice.Editable()
	docOrder := rOrder.Editable()

	//	priceOne, _ := strconv.Atoi(arr[0]["price"])
	//	priceTwo, _ := strconv.Atoi(arr[1]["price"])
	//////	priceThree, _ := strconv.Atoi(arr[2]["price"])

	// totalPrice := priceOne + priceTwo + priceThree

	// Заданная дата
	dateStr := dateText
	layout := "02.01.2006"

	// Парсим дату
	currentDate, err := time.Parse(layout, dateStr)
	if err != nil {
		fmt.Println("Ошибка парсинга даты:", err)
		return
	}

	// Вычисляем первый день предыдущего месяца
	firstDayPrevMonth := currentDate.AddDate(0, -1, -currentDate.Day()+1)

	// Вычисляем последний день предыдущего месяца
	lastDayPrevMonth := firstDayPrevMonth.AddDate(0, 1, -1)

	// Вывод результатов
	fmt.Println("Первый день предыдущего месяца:", firstDayPrevMonth.Format(layout))
	fmt.Println("Последний день предыдущего месяца:", lastDayPrevMonth.Format(layout))

	// Переменные для замены
	variables := map[string]string{
		"Number":                    "1",
		"DateDocument":              dateText,
		"DescriptionFirstPosition":  arr[0]["description"],
		"DescriptionSecondPosition": arr[1]["description"],
		"DescriptionThirdPosition":  arr[2]["description"],
		"PriceFirstPosition":        arr[0]["price"],
		"PriceSecondPosition":       arr[1]["price"],
		"PriceThirdPosition":        arr[2]["price"],
		"TotalPrice":                arr[3]["total_price"],
		"TextPrice":                 arr[3]["total_price_text"],
		"DatePrevFirst":             firstDayPrevMonth.Format(layout),
		"DatePrevLast":              lastDayPrevMonth.Format(layout),
	}

	// Заменяем переменные в тексте
	for key, value := range variables {
		placeholder := fmt.Sprintf("%s", key)

		err := docAct.Replace(placeholder, value, -1)
		if err != nil {
			fmt.Println(err)
			return
		}

		fmt.Printf("Проверка текста: %s\n", key)
		fmt.Printf("Заменяем %s на %s\n", placeholder, value)
	}

	// Заменяем переменные в тексте
	for key, value := range variables {
		placeholder := fmt.Sprintf("%s", key)

		err := docInvoice.Replace(placeholder, value, -1)
		if err != nil {
			fmt.Println(err)
			return
		}

		fmt.Printf("Проверка текста: %s\n", key)
		fmt.Printf("Заменяем %s на %s\n", placeholder, value)
	}

	// Заменяем переменные в тексте
	for key, value := range variables {
		placeholder := fmt.Sprintf("%s", key)

		err := docOrder.Replace(placeholder, value, -1)
		if err != nil {
			fmt.Println(err)
			return
		}

		fmt.Printf("Проверка текста: %s\n", key)
		fmt.Printf("Заменяем %s на %s\n", placeholder, value)
	}

	os.Mkdir("./dist/"+dateText, 0777)

	// Сохраняем измененный документ
	outputFileAct := "./dist/" + dateText + "/act.docx"
	err = docAct.WriteToFile(outputFileAct)
	if err != nil {
		log.Fatalf("Ошибка сохранения документа: %v", err)
	}

	// Сохраняем измененный документ
	outputFileInvoice := "./dist/" + dateText + "/invoice.docx"
	err = docInvoice.WriteToFile(outputFileInvoice)
	if err != nil {
		log.Fatalf("Ошибка сохранения документа: %v", err)
	}

	// Сохраняем измененный документ
	outputFileOrder := "./dist/" + dateText + "/order.docx"
	err = docOrder.WriteToFile(outputFileOrder)
	if err != nil {
		log.Fatalf("Ошибка сохранения документа: %v", err)
	}

	fmt.Printf("Документ успешно сохранен: %s\n", outputFileAct)
	fmt.Printf("Документ успешно сохранен: %s\n", outputFileInvoice)
	fmt.Printf("Документ успешно сохранен: %s\n", outputFileOrder)
}
