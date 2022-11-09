//********************************************************************************
// БЛОК 0
// Ініціація Задачі

// Глобальні змінні задачі: ОСНОВНІ
var chData = {
	'chDate': {},						// Дата за яку обробляються дані 
	'iName': 0, 'NName': 0,			// Лічильник та кількість завантажених файлів
	'chCSV': {							// Хеш даних CSV телеметрії:
		'frequency': [], 				// - частоти
		'state': [], 					// - активацій
	},
	'chXLS': {},						// Хеш даних XLSX телеметрії
	'arBlock': {},						// Кількість блоків по годинам
	//'chJSON': {},						// Хеш даних XLSX телеметрії
	'iRecord': 0, 						// Кількість записів по: 
	'chSort': {},						// Розсортовані дані
	'arAnalys': [],					// Проаналізовані дані
	'arExcel': []						// Масив даних для формування та вивантаження EXCEL
};

// Глобальні обєкти задачі
let elem = document.documentElement.childNodes;		// БЛОК 1.1 Блокуванна завантаження фалів елементами сторінки
let items = document.querySelectorAll('.upload');	// БЛОК 1.2 Розблукування поля завантаження файлів даних


//********************************************************************************
// БЛОК 1
// Вивід інтерфейсу та завантаження даних

// БЛОК 1.1
// Інтерфейс - завнтаження drag and drop
// html, body - не реагувати на пертягування та скидання файлів
function handleDragOverDropBody(event) { event.preventDefault(); event.stopPropagation(); return false; }
for (let i = 0; i < elem.length; i++) { // перебір елементів ТІЛЬКИ за допомогою for
	elem[i].addEventListener('dragover', handleDragOverDropBody);
	elem[i].addEventListener('drop', handleDragOverDropBody);
}

// БЛОК 1.2
// Завантаження фалів
// div - отримання пертягнутих файлів
function handleDragOver() { this.classList.add('over'); return false; }
function handleDragLeave() { this.classList.remove('over'); return false; }
function handleDrop(event) {
	this.classList.remove('over'); event.preventDefault();
	let files = event.dataTransfer.files;
	Upload(files);
}
// Привязка функцій до поля завантаження файлів для обробки подій: наведення/відведення фокусу, скидання 
items.forEach(function (item) {
	item.addEventListener('dragover', handleDragOver);
	item.addEventListener('dragleave', handleDragLeave);
	item.addEventListener('drop', handleDrop);
});
//Прогрес завантаження файлів: показати		
function ShowProgress(Percent) {
	progres_bar.style.width = Percent + '%';
	if (Percent >= 100) setTimeout(HideProgress, 1000);
}
//Прогрес завантаження файлів: приховати
function HideProgress() {
	progres_bar.style.width = '0%';
}

// БЛОК 1.3
// Читання фалів та консолідація даних

// Встановлення початкових значень та стилів
function setDefaltStyle(Choise, Parametr) {
	// Встановлення початкових значень та стилів ПЕРЕД ЗАВАНТАЖЕННЯМ  
	if (Choise == 0) {
		// Змінні
		chData = {
			'chDate': {},						// Дата за яку обробляються дані 
			'iName': 0, 'NName': Parametr,			// Лічильник та кількість завантажених файлів
			'chCSV': {							// Хеш даних CSV телеметрії:
				'frequency': [], 				// - частоти
				'state': [], 					// - активацій
			},
			//'chJSON': {},						// Хеш даних XLSX телеметрії
			'chXLS': {},						// Хеш даних XLSX телеметрії
			'arBlock': {},						// Кількість блоків по годинам
			'iRecord': 0, 						// Кількість записів по: 
			'chSort': {},						// Розсортовані дані
			'arAnalys': [],					// Проаналізовані дані
			'arExcel': []						// Масив даних для формування та вивантаження EXCEL
		};

		// Стилі
		process.style.visibility = 'hidden';														// Приховати кнопку вивантаження xlsx-файлу аналізу телеметрії 
	}
	//Встановлення значень та стилів ПІСЛЯ ЗАВАНТАЖЕННЯМ  
	if (Choise == 1) {
		process.style.visibility = 'visible';														// Відобразити кнопку вивантаження xlsx-файлу аналізу телеметрії 
	}
}
//Встановлення дати та часу за Киїівським часом
function setDateTimeKiev(strDateTimeUTS) {
	//Прибираємо зайві символи
	let strDTUTS = strDateTimeUTS.split('+')[0], strDTuts = strDTUTS.split('.')[0];
	let DTKiev = new Date(Date.parse(strDTuts));									// Перехід на Київський час

	//Визначаємо зимній/літній час (+2/+3 години відповідно)
	let valOffSet = DTKiev.getTimezoneOffset(), WinSum = Math.abs(valOffSet / 60);		// - 2 години так вони вже додані при перході на Києвський час   
	DTKiev = new Date(Date.parse(strDTuts) + WinSum * 3600 * 1000);

	//Виділяємо Рік Місяць День, Годину, Хвилину, Секунду
	let YY = DTKiev.getFullYear(), M = '0' + (DTKiev.getMonth() + 1), D = '0' + DTKiev.getDate(), h = '0' + DTKiev.getHours(), m = '0' + DTKiev.getMinutes(), s = '0' + DTKiev.getSeconds();
	let s10 = Math.trunc(s / 10) * 10, ss10 = '00'; if (s10 >= 10) { ss10 = s10; }
	let MM = M.substr(-2, 2), DD = D.substr(-2, 2), hh = h.substr(-2, 2), mm = m.substr(-2, 2), ss = s.substr(-2, 2);

	let dateD = YY + '-' + MM + '-' + DD, dateC = DD + '.' + MM + '.' + YY;
	let time60 = hh + ':' + mm + ':' + ss, time10 = hh + ':' + mm + ':' + ss10, time00 = hh + ':' + mm + ':00';
	let chDateTimeKiev = { 'dateD': dateD, 'dateC': dateC, 'time60': time60, 'time10': time10, 'time00': time00, 'hh': hh };
	return chDateTimeKiev;
}
// Отримання дати з назви файлу
function getCHDateCSV(filename) {
	let aD = filename.split('_')[0].split('.'), YY = '20' + aD[0], MM = aD[1], DD = aD[2];
	let dateD = YY + '-' + MM + '-' + DD, dateC = DD + '.' + MM + '.' + YY, dateP = YY + '.' + MM + '.' + DD;
	let chDate = { 'dateD': dateD, 'dateC': dateC, 'dateP': dateP };
	return chDate;
}
// Додавання ПОРОЖНЬОГО ХЕШУ
function setNULLArrayCSV(Type, dateD, hh, keyTPP, time60) {
	if (Type == 'frequency') {
		if (!chData.chCSV[Type][dateD]) { chData.chCSV[Type][dateD] = {}; }
		if (!chData.chCSV[Type][dateD][hh]) { chData.chCSV[Type][dateD][hh] = {}; }
		if (!chData.chCSV[Type][dateD][hh][keyTPP]) { chData.chCSV[Type][dateD][hh][keyTPP] = {}; }
		if (!chData.chCSV[Type][dateD][hh][keyTPP][time60]) { chData.chCSV[Type][dateD][hh][keyTPP][time60] = {}; }
	}
	if (Type == 'state') {
		if (!chData.chCSV[Type][dateD]) { chData.chCSV[Type][dateD] = {}; }
		if (!chData.chCSV[Type][dateD][keyTPP]) { chData.chCSV.state[dateD][keyTPP] = {}; }
	}
	if (Type == 'analys') {
		if (!chData.arAnalys[dateD]) { chData.arAnalys[dateD] = {}; }
		if (!chData.arAnalys[dateD][hh]) {
			chData.arAnalys[dateD][hh] = { '50.01': { 0: { 'mid': 0, 'total': 0, 'count': 0 } }, '49.99': { 0: { 'mid': 0, 'total': 0, 'count': 0 } }, 'minmax': { 'min': 50, 'max': 50 } };
		}
	}
}

// Читаємо дані з EXCEL та зберігаємо у хеш chData.chCSV-дані
function readFileCSV(file) {
	let fr = new FileReader();																				// EXCEL - порожній обєкт файлу
	chData.chDate = getCHDateCSV(file.name);																// EXCEL - назва файлу - fileName = file.name

	fr.onload = function () {
		var arData = fr.result.split(/\r\n|\n/);														// Масив даних
		for (let i = 1; i < arData.length; i++) {
			// Парсинг та отримання даних для одного запису
			let arVal = arData[i].split(','), DateTimeUTS = arVal[0], Value = arVal[1], TPP = arVal[2], Block = arVal[3], Type = arVal[4];
			// Визначення дати та часу за Киїівським часом
			let chDTK = setDateTimeKiev(DateTimeUTS), dateD = chDTK.dateD, dateC = chDTK.dateC, hh = chDTK.hh, time60 = chDTK.time60, time10 = chDTK.time10, time00 = chDTK.time00;

			let keyTPP = TPP + "_" + Block;																// Формуємо ключ для хешу АКТИВАЦІЙ
			if (dateD != 'NaN-N1-aN' && chData.chDate.dateD == dateD) {							// Первірка на порожню дату та даним з датою, що відповідають даті файлу

				if (Type == 'frequency') {
					setNULLArrayCSV(Type, dateD, hh, keyTPP, time60);									// Додавання ПОРОЖНЬОГО ХЕШУ до даних ЧАСТОТИ
					setNULLArrayCSV('analys', dateD, hh, keyTPP, time60);								// Додавання ПОРОЖНЬОГО ХЕШУ до даних АНАЛІЗУ
					chData.chCSV.frequency[dateD][hh][keyTPP][time60] = Value;					// Запис ЧАСТОТИ //{ 'frequency': Value, 'dateC': dateC, 'time10': time10, 'time00': time00 };
				}
				if (Type == 'state') {
					setNULLArrayCSV(Type, dateD, hh, keyTPP, time60);									// Додавання ПОРОЖНЬОГО ХЕШУ до даних АКТИВАЦІЙ
					chData.chCSV.state[dateD][keyTPP][time60] = Value;							// Запис АКТИВАЦІЇ
				}
			}
		}
		chData.iRecord = chData.iRecord + arData.length;									// Кількість записів: всі
		upload_count.innerHTML = chData.iRecord;												// Вивід кількості записів: всі
		chData.iName++;																				// Лічильник файлів: всі
		let Percent = Math.round((chData.iName / chData.NName) * 100);					// Процент завантаження файлів
		ShowProgress(Percent);																		// Прогрес завантаження файлів		
	};
	fr.readAsBinaryString(file);																	// Запуск читання файлу після успішного завантаження
}



// Отримання дати з назви файлу
function getCHDateXLS(filename) {
	let aF = filename.split('_'), aD = aF[0].split('.'), YY = aD[0], MM = aD[1], DD = aD[2], TPP = aF[1];
	let dateD = YY + '-' + MM + '-' + DD, dateC = DD + '.' + MM + '.' + YY, dateP = YY + '.' + MM + '.' + DD;
	let chDate = { 'dateD': dateD, 'dateC': dateC, 'dateP': dateP, 'TPP': TPP };
	return chDate;
}
// Додавання ПОРОЖНЬОГО ХЕШУ
//function setNULLArrayJSON(dateD, keyTPP) {
//	if (!chData.chJSON[dateD]) { chData.chJSON[dateD] = {}; }
//	if (!chData.chJSON[dateD][keyTPP]) { chData.chJSON[dateD][keyTPP] = {}; }
//}
// Додавання ПОРОЖНЬОГО ХЕШУ
function setNULLArrayXLS(dateD, hh, keyTPP, time60) {
	if (!chData.chXLS[dateD]) { chData.chXLS[dateD] = {}; }
	if (!chData.chXLS[dateD][hh]) { chData.chXLS[dateD][hh] = {}; }
	if (!chData.chXLS[dateD][hh][keyTPP]) { chData.chXLS[dateD][hh][keyTPP] = {}; }
	if (!chData.chXLS[dateD][hh][keyTPP][time60]) { chData.chXLS[dateD][hh][keyTPP][time60] = { 'time60': '-', 'frequency': 0, 'state': 0 }; }
}

//Визначення часу та години
//function setDateTimeXLS(timeXLS, iSecond) {
//	time60=getTimeXLS(iSecond);
//	let past = 1 / 86399, second = Math.round(timeXLS / past);
//	let h = Math.trunc(second / 3600), h0 = '0' + h, hh = h0.substr(-2, 2);
//	let m = Math.trunc((second - h * 3600) / 60), m0 = '0' + m, mm = m0.substr(-2, 2);
//	let s = second - h * 3600 - m * 60, s0 = '0' + s, ss = s0.substr(-2, 2);
//	let time60 = hh + ':' + mm + ':' + ss, chTime = { 'time60': time60, 'hh': hh };
//	return chTime;
//}

// Читаємо дані з EXCEL та зберігаємо у хеш chData.chName-назви файлів та chData.chCSV-дані
function readFileXLS(file) {
	let fr = new FileReader();																					// EXCEL - порожній обєкт файлу
	chData.chDate = getCHDateXLS(file.name);																// EXCEL - назва файлу - fileName = file.name
	let dateD = chData.chDate.dateD, TPP = chData.chDate.TPP;										// Получаем Дату, название ПДП 
	fr.onload = function () {
		let iN = chData.iName;																					// Лічильник завантажених фалів

		let data = fr.result;																					// EXCEL - присвоєння обєкту завантаженого файлу
		let workbook = XLSX.read(data, { type: 'binary' });											// EXCEL - читаємо вкладки
		workbook.SheetNames.forEach(sheet => {																// EXCEL - перебір вкладок 
			let Block = sheet, keyTPP = TPP + '_' + Block;												// Назва вкладки - номер блоку
			let arExcel = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]); 		// ХЕШ ДАНИХ - Записуємо нові дані
			for (let isecond = 0; isecond < arExcel.length; isecond++) {
				let chTime = getTimeXLS(isecond), time60 = chTime.time60, hh = chTime.hh, h = chTime.h; //timeXLS = arExcel[isecond]['Время'],
				let frequency = arExcel[isecond]['Frequency'];
				let state = arExcel[isecond]['State'];
				setNULLArrayCSV('analys', dateD, h, keyTPP, time60);									// Додавання ПОРОЖНЬОГО ХЕШУ до даних АНАЛІЗУ
				setNULLArrayXLS(dateD, h, keyTPP, isecond);
				chData.chXLS[dateD][h][keyTPP][isecond] = { 'time60': time60, 'frequency': frequency, 'state': state };
			}
			chData.iRecord += arExcel.length;																// Кількість записів: всі
			upload_count.innerHTML = chData.chDate.dateD + ' - ' + chData.iRecord;				// Виводимо кількість записів

			//setNULLArrayJSON(dateD, keyTPP);
			//chData.chJSON[dateD][keyTPP] = arExcel; 													// ХЕШ ДАНИХ - Записуємо нові дані

		});
		chData.iName++;																							// Лічільник завантажених файлів
		let Percent = Math.round((chData.iName / chData.NName) * 100);								// Процент завантаження файлів
		ShowProgress(Percent);																					// Прогрес завантаження файлів		
	};
	fr.readAsBinaryString(file);																				// Тип читання - читання як двійкова строка
}

//Завантаження на сервер та перебір отриманих файлів
function Upload(files) {
	setDefaltStyle(0, files.length);																			//	Встановлення початкових значень та стилів ПЕРЕД ЗАВАНТАЖЕННЯМ  
	for (let i = 0; i < chData.NName; i++) {
		//readFileCSV(files[i]);																				// Читання фалів
		readFileXLS(files[i]);																					// Читання фалів
	}
}

//********************************************************************************
// БЛОК 2
// Аналіз завантажених даних

// Інкремент часу
function IcremetTime(dateD, time60, sec) {
	let tm60 = new Date(Date.parse(dateD + ' ' + time60) + sec * 1000);
	let h = '0' + tm60.getHours(), m = '0' + tm60.getMinutes(), s = '0' + tm60.getSeconds();
	let hh = h.substr(-2, 2), mm = m.substr(-2, 2), ss = s.substr(-2, 2);
	time60 = hh + ':' + mm + ':' + ss;
	return time60;
}

// Пошук АКТИВАЦІЇ за датою/часом та ПДП
function setState(dateD, keyTPP, time60) {
	let set = 'no', iEnd = 0, val = 0, time = time60;
	while (set == 'no') {
		if (!chData.chCSV.state[dateD][keyTPP][time60]) { time60 = IcremetTime(dateD, time60, 1); }		// інкремент часу на 1 секунду
		else { set = 'yes'; val = chData.chCSV.state[dateD][keyTPP][time60]; }									// значення АКТИВАЦІЇ знайдено 
		if (iEnd > 60) { set = 'yes'; }																						// Передчасне завершення пошуку на проміжку 1 хвилина
		iEnd++;
	}
	let state = val;
	return state
}
//Визначення проміжку часу
function betweenTimeCSV(dateD, beginT, endT) {
	let DTbegin = Date.parse(dateD + ' ' + beginT), DTend = Date.parse(dateD + ' ' + endT), T;
	if (DTbegin > DTend) {
		T = DTbegin; DTbegin = DTend; DTend = T;
		T = beginT; beginT = endT; endT = T;
	}
	if (DTbegin == DTend) { DTend += 1000; }					// У разі співпадіння початку та кінця інтервалу 
	let beSecond = (DTend - DTbegin) / 1000, beH = Math.trunc(beSecond / 60), beS = beSecond - beH * 60, beTime = beH + ':' + beS;

	let chbetweenTimeCSV = { 'beginT': beginT, 'endT': endT, 'beTime': beTime, 'beSecond': beSecond };
	return chbetweenTimeCSV;
}

function analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime) {
	let beSecond = chBeTime.beSecond, beginT = chBeTime.beginT, endT = chBeTime.endT, beTime = chBeTime.beTime;

	chData.arAnalys[dateD][hh][direct][0].total += beSecond;	// Запис інтервалу виходу за 					49.99
	chData.arAnalys[dateD][hh][direct][0].count++;								// Запис кількості інтервалів виходу за 	49.99
	// Визначення середнього інтервалу виходу за 49.99
	let total = chData.arAnalys[dateD][hh][direct][0].total, count = chData.arAnalys[dateD][hh][direct][0].count;
	chData.arAnalys[dateD][hh][direct][0].mid = total / count;

	chData.arAnalys[dateD][hh][direct][count] = { 'total': beSecond, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
}

// Аналіз даних
function analiseDataCSV() {
	// Аналіз даних
	let frequency, state, direct = '50', beginT = '-', endT = '-', chBeTime = {}, Max, Min, timeEnd;
	for (let dateD in chData.chCSV.frequency) {
		for (let hh in chData.chCSV.frequency[dateD]) {
			for (let keyTPP in chData.chCSV.frequency[dateD][hh]) {

				Max = 50; Min = 50;
				direct = '50';																// Встановлення напрямку				50
				beginT = '-';																// Скидання часу ПОЧАТКУ інтервалу

				for (let time60 in chData.chCSV.frequency[dateD][hh][keyTPP]) {

					state = setState(dateD, keyTPP, time60); frequency = chData.chCSV.frequency[dateD][hh][keyTPP][time60];
					if (state == 1) {
						if (Min > frequency && frequency > 45) { Min = frequency; }
						if (Max < frequency) { Max = frequency; }
						if (frequency > 50.01) {												// Вихід за мертву зону 				50.01 
							if (beginT != '-' && direct == '49.99') {						// Якщо встановлена час початку

								endT = time60;														// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeCSV(dateD, beginT, endT);				// Отрмання інтервалу виходу за  	49.99 
								analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime);	// Оновлення 0-го та створення N-го запису результату аналізу	
							}
							direct = '50.01';														// Встановлення напрямку				50.01
							beginT = time60;														// Встановлення часу ПОЧАТКУ інтервалу
						}
						if (frequency < 49.99) {												// Вихід за мертву зону 				49.99
							if (beginT != '-' && direct == '50.01') {						// Якщо встановлена час початку

								endT = time60;														// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeCSV(dateD, beginT, endT);				// Отрмання інтервалу виходу за 		50.01
								analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime);	// Оновлення 0-го та створення N-го запису результату аналізу
							}
							direct = '49.99';															// Встановлення напрямку				49.99
							beginT = time60;														// Встановлення часу ПОЧАТКУ інтервалу
						}
						if (frequency >= 49.99 && frequency <= 50.01) {					//Повернення у мертву зону 			50.01 або 49.99
							if (beginT != '-' && direct != '50') {							// Якщо встановлена час початку

								endT = time60;														// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeCSV(dateD, beginT, endT);				// Отрмання інтервалу виходу за	 	50.01 або 49.99
								analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime);	// Оновлення 0-го та створення N-го запису результату аналізу
							}
							direct = '50';															// Встановлення напрямку				50
							beginT = '-';															// Скидання часу ПОЧАТКУ інтервалу
						}
					}
					else {
						if (beginT != '-' && direct != '50') {								// Якщо встановлена час початку

							endT = time60;															// Встановлення часу ЗАВЕРШЕННЯ інтервалу
							chBeTime = betweenTimeCSV(dateD, beginT, endT);					// Інтервал виходу за мертву зону 	50.01 або 49.99
							analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime);		// Оновлення 0-го та створення N-го запису результату аналізу
						}
						direct = '50';																// Встановлення напрямку				50
						beginT = '-';																// Скидання часу ПОЧАТКУ інтервалу

					}
					timeEnd = time60;
				}
				// Обробка останньог інтервалу виходу за мертву зону
				if (beginT != '-' && direct != '50') {								// Якщо встановлена час початку

					endT = timeEnd;															// Встановлення часу ЗАВЕРШЕННЯ інтервалу
					chBeTime = betweenTimeCSV(dateD, beginT, endT);					// Інтервал виходу за мертву зону 	50.01 або 49.99
					analiseRecordCSV(dateD, hh, direct, keyTPP, chBeTime);		// Оновлення 0-го та створення N-го запису результату аналізу
				}

				let max = chData.arAnalys[dateD][hh]['minmax'].max, min = chData.arAnalys[dateD][hh]['minmax'].min;
				if (min > Min) { chData.arAnalys[dateD][hh]['minmax'].min = Min; }		// МІНІМАЛЬНА частота
				if (max < Max) { chData.arAnalys[dateD][hh]['minmax'].max = Max; }		// МАКСИМАЛЬНА частота	
			}
		}
	}
	//Створення масиву для створення Excel файлу
	let ie = 0, total, keyTPP, beTime, beH, beS, numBlock = chData.NName / 2, dateP = chData.chDate.dateP;
	for (let dateD in chData.arAnalys) {
		for (let hh in chData.arAnalys[dateD]) {
			//49.99
			let midlow = Math.round(chData.arAnalys[dateD][hh]['49.99'][0].mid); //'total': beSecond,'beTime':beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP 
			beH = Math.trunc(midlow / 60); beS = midlow - beH * 60;
			let beTimeLow = beH + ' minutes ' + beS + ' seconds';

			let totallow = chData.arAnalys[dateD][hh]['49.99'][0].total;
			beH = Math.trunc((totallow / numBlock) / 60); beS = Math.trunc(totallow / numBlock) - beH * 60;
			let beTimeBlockLow = beH + ' minutes ' + beS + ' seconds';

			let countlow = chData.arAnalys[dateD][hh]['49.99'][0].count;
			let chLow = {};
			for (let i in chData.arAnalys[dateD][hh]['49.99']) {
				if (i > 0) {
					total = chData.arAnalys[dateD][hh]['49.99'][i].total;
					beTime = chData.arAnalys[dateD][hh]['49.99'][i].beTime;
					beginT = chData.arAnalys[dateD][hh]['49.99'][i].beginT;
					endT = chData.arAnalys[dateD][hh]['49.99'][i].endT;
					keyTPP = chData.arAnalys[dateD][hh]['49.99'][i].keyTPP;
					chLow[i - 1] = { 'total': total, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
				}
			}
			let stLow = JSON.stringify(chLow);

			//50.01
			let midhi = Math.round(chData.arAnalys[dateD][hh]['50.01'][0].mid);
			beH = Math.trunc(midhi / 60); beS = midhi - beH * 60;
			let beTimeHi = beH + ' minutes ' + beS + ' seconds';

			let totalhi = chData.arAnalys[dateD][hh]['50.01'][0].total;
			beH = Math.trunc((totalhi / numBlock) / 60); beS = Math.trunc(totalhi / numBlock) - beH * 60;
			let beTimeBlockHi = beH + ' minutes ' + beS + ' seconds';

			let counthi = chData.arAnalys[dateD][hh]['50.01'][0].count;

			let chHi = {};
			for (let i in chData.arAnalys[dateD][hh]['50.01']) {
				if (i > 0) {
					total = chData.arAnalys[dateD][hh]['50.01'][i].total;
					beTime = chData.arAnalys[dateD][hh]['50.01'][i].beTime;
					beginT = chData.arAnalys[dateD][hh]['50.01'][i].beginT;
					endT = chData.arAnalys[dateD][hh]['50.01'][i].endT;
					keyTPP = chData.arAnalys[dateD][hh]['50.01'][i].keyTPP;
					chHi[i - 1] = { 'total': total, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
				}
			}
			let stHi = JSON.stringify(chHi);

			let max = chData.arAnalys[dateD][hh].minmax.max;
			let min = chData.arAnalys[dateD][hh].minmax.min;

			//chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '<49.99', 'Duration': midlow + ' -> ' + beTimeLow, 'minmax': min, 'json': stLow }; ie++;
			//chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '>50.01', 'Duration': midhi + ' -> ' + beTimeHi, 'minmax': max, 'json': stHi }; ie++;
			//chData.arExcel[ie] = { 'Date': dateP + ' ' + dateD, 'Hour': hh, 'Direction': '<49.99', 'Duration': beTimeLow, 'MinMax': min }; ie++;
			//chData.arExcel[ie] = { 'Date': dateP + ' ' + dateD, 'Hour': hh, 'Direction': '>50.01', 'Duration': beTimeHi, 'MinMax': max }; ie++;

			//chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '<49.99', 'Duration': beTimeLow, 'MinMax': min }; ie++;
			//chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '>50.01', 'Duration': beTimeHi, 'MinMax': max }; ie++;

			chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '<49.99', 'Duration': beTimeBlockLow, 'MinMax': min }; ie++;
			chData.arExcel[ie] = { 'Date': dateP, 'Hour': hh, 'Direction': '>50.01', 'Duration': beTimeBlockHi, 'MinMax': max }; ie++;
		}
	}

	setDefaltStyle(1, 0); // Візуалізація кнопки завантаження
}
//Визначення проміжку часу
function getTimeXLS(Second) {
	let h = Math.trunc(Second / 3600), m = Math.trunc((Second - h * 3600) / 60), s = Second - (h * 3600 + m * 60);
	let h0 = '0' + h, m0 = '0' + m, s0 = '0' + s;
	let hh = h0.substr(-2, 2), mm = m0.substr(-2, 2), ss = s0.substr(-2, 2);
	let time60 = hh + ':' + mm + ':' + ss; chTime = { 'time60': time60, 'hh': hh, 'h': h };
	return chTime;
}
//Визначення проміжку часу
function betweenTimeXLS(beginS, endS) {
	let beginT = getTimeXLS(beginS).time60, endT = getTimeXLS(endS).time60, beSecond = endS - beginS, beTime = getTimeXLS(beSecond).time60;

	let chbetweenTimeCSV = { 'beginT': beginT, 'endT': endT, 'beTime': beTime, 'beSecond': beSecond };
	return chbetweenTimeCSV;
}
function analiseRecordXLS(dateD, hh, direct, keyTPP, chBeTime) {
	let beSecond = chBeTime.beSecond, beginT = chBeTime.beginT, endT = chBeTime.endT, beTime = chBeTime.beTime;

	chData.arAnalys[dateD][hh][direct][0].total += beSecond;	// Запис інтервалу виходу за 					49.99
	chData.arAnalys[dateD][hh][direct][0].count++;								// Запис кількості інтервалів виходу за 	49.99
	// Визначення середнього інтервалу виходу за 49.99
	let total = chData.arAnalys[dateD][hh][direct][0].total, count = chData.arAnalys[dateD][hh][direct][0].count;
	chData.arAnalys[dateD][hh][direct][0].mid = total / count;

	chData.arAnalys[dateD][hh][direct][count] = { 'total': beSecond, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
}
// Аналіз даних
function analiseDataXLS() {
	// Аналіз даних
	let frequency, state, direct = '50', beginT = '-', endT = '-', chBeTime = {}, Max, Min, timeEnd, beginS = '-', endS = '-', time60;
	for (let dateD in chData.chXLS) {
		for (let h in chData.chXLS[dateD]) {
			for (let keyTPP in chData.chXLS[dateD][h]) {

				Max = 50; Min = 50;
				direct = '50';																		// Встановлення напрямку				50
				beginS = '-';																		// Скидання часу ПОЧАТКУ інтервалу
				if (!chData.arBlock[h]) { chData.arBlock[h] = 1; } chData.arBlock[h]++;					// Кількість блоків по годинамnumBlock++;																	// Кількість блоків
				timeEnd = 0;
				for (let isecond in chData.chXLS[dateD][h][keyTPP]) {

					state = chData.chXLS[dateD][h][keyTPP][isecond].state;
					frequency = chData.chXLS[dateD][h][keyTPP][isecond].frequency;
					time60 = chData.chXLS[dateD][h][keyTPP][isecond].time60;


					if (state == 1) {
						if (Min > frequency && frequency > 45) { Min = frequency; }
						if (Max < frequency) { Max = frequency; }
						if (frequency > 50.01) {												// Вихід за мертву зону 				50.01 
							if (beginT != '-' && direct == '49.99') {						// Якщо встановлена час початку

								endS = isecond;														// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeXLS(beginS, endS);					// Отрмання інтервалу виходу за  	49.99 
								analiseRecordXLS(dateD, h, direct, keyTPP, chBeTime);// Оновлення 0-го та створення N-го запису результату аналізу
								beginS = '-';
							}
							direct = '50.01';														// Встановлення напрямку				50.01
							if (beginS == '-') { beginS = isecond; }						// Встановлення часу ПОЧАТКУ інтервалу
						}
						if (frequency < 49.99) {												// Вихід за мертву зону 				49.99
							if (beginS != '-' && direct == '50.01') {						// Якщо встановлена час початку

								endS = isecond;													// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeXLS(beginS, endS);					// Отрмання інтервалу виходу за 		50.01
								analiseRecordXLS(dateD, h, direct, keyTPP, chBeTime);	// Оновлення 0-го та створення N-го запису результату аналізу
								beginS == '-';
							}
							direct = '49.99';														// Встановлення напрямку				49.99
							if (beginS == '-') { beginS = isecond; }						// Встановлення часу ПОЧАТКУ інтервалу
						}
						if (frequency >= 49.99 && frequency <= 50.01) {					//Повернення у мертву зону 			50.01 або 49.99
							if (beginS != '-' && direct != '50') {							// Якщо встановлена час початку

								endS = isecond;													// Встановлення часу ЗАВЕРШЕННЯ інтервалу
								chBeTime = betweenTimeXLS(beginS, endS);					// Отрмання інтервалу виходу за	 	50.01 або 49.99
								analiseRecordXLS(dateD, h, direct, keyTPP, chBeTime);// Оновлення 0-го та створення N-го запису результату аналізу
							}
							direct = '50';															// Встановлення напрямку				50
							beginS = '-';															// Скидання часу ПОЧАТКУ інтервалу
						}
					}
					else {
						if (beginS != '-' && direct != '50') {								// Якщо встановлена час початку

							endS = isecond;														// Встановлення часу ЗАВЕРШЕННЯ інтервалу
							chBeTime = betweenTimeXLS(beginS, endS);						// Інтервал виходу за мертву зону 	50.01 або 49.99
							analiseRecordXLS(dateD, h, direct, keyTPP, chBeTime);	// Оновлення 0-го та створення N-го запису результату аналізу
						}
						direct = '50';																// Встановлення напрямку				50
						beginS = '-';																// Скидання часу ПОЧАТКУ інтервалу

					}
					timeEnd = isecond;
				}
				// Обробка останньог інтервалу виходу за мертву зону
				if (beginS != '-' && direct != '50') {										// Якщо встановлена час початку

					endS = timeEnd;																// Встановлення часу ЗАВЕРШЕННЯ інтервалу
					chBeTime = betweenTimeXLS(beginS, endS);								// Інтервал виходу за мертву зону 	50.01 або 49.99
					analiseRecordXLS(dateD, h, direct, keyTPP, chBeTime);			// Оновлення 0-го та створення N-го запису результату аналізу
				}

				let max = chData.arAnalys[dateD][h]['minmax'].max, min = chData.arAnalys[dateD][h]['minmax'].min;
				if (min > Min) { chData.arAnalys[dateD][h]['minmax'].min = Min; }		// МІНІМАЛЬНА частота
				if (max < Max) { chData.arAnalys[dateD][h]['minmax'].max = Max; }		// МАКСИМАЛЬНА частота	
			}
		}
	}
	//Створення масиву для створення Excel файлу
	let ie = 0, total, keyTPP, beTime, beH, beS, dateP = chData.chDate.dateP;
	for (let dateD in chData.arAnalys) {
		for (let hh in chData.arAnalys[dateD]) {
			//49.99
			let midlow = Math.round(chData.arAnalys[dateD][hh]['49.99'][0].mid); //'total': beSecond,'beTime':beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP 
			beH = Math.trunc(midlow / 60); beS = midlow - beH * 60;
			let beTimeLow = beH + ' minutes ' + beS + ' seconds';

			let totallow = chData.arAnalys[dateD][hh]['49.99'][0].total;
			beH = Math.trunc((totallow / chData.arBlock[hh]) / 60); beS = Math.trunc(totallow / chData.arBlock[hh]) - beH * 60;
			let beTimeBlockLow = beH + ' minutes ' + beS + ' seconds';

			let countlow = chData.arAnalys[dateD][hh]['49.99'][0].count;
			let chLow = {};
			for (let i in chData.arAnalys[dateD][hh]['49.99']) {
				if (i > 0) {
					total = chData.arAnalys[dateD][hh]['49.99'][i].total;
					beTime = chData.arAnalys[dateD][hh]['49.99'][i].beTime;
					beginT = chData.arAnalys[dateD][hh]['49.99'][i].beginT;
					endT = chData.arAnalys[dateD][hh]['49.99'][i].endT;
					keyTPP = chData.arAnalys[dateD][hh]['49.99'][i].keyTPP;
					chLow[i - 1] = { 'total': total, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
				}
			}
			let stLow = JSON.stringify(chLow);

			//50.01
			let midhi = Math.round(chData.arAnalys[dateD][hh]['50.01'][0].mid);
			beH = Math.trunc(midhi / 60); beS = midhi - beH * 60;
			let beTimeHi = beH + ' minutes ' + beS + ' seconds';

			let totalhi = chData.arAnalys[dateD][hh]['50.01'][0].total;
			beH = Math.trunc((totalhi / chData.arBlock[hh]) / 60); beS = Math.trunc(totalhi / chData.arBlock[hh]) - beH * 60;
			let beTimeBlockHi = beH + ' minutes ' + beS + ' seconds';

			let counthi = chData.arAnalys[dateD][hh]['50.01'][0].count;

			let chHi = {};
			for (let i in chData.arAnalys[dateD][hh]['50.01']) {
				if (i > 0) {
					total = chData.arAnalys[dateD][hh]['50.01'][i].total;
					beTime = chData.arAnalys[dateD][hh]['50.01'][i].beTime;
					beginT = chData.arAnalys[dateD][hh]['50.01'][i].beginT;
					endT = chData.arAnalys[dateD][hh]['50.01'][i].endT;
					keyTPP = chData.arAnalys[dateD][hh]['50.01'][i].keyTPP;
					chHi[i - 1] = { 'total': total, 'beTime': beTime, 'beginT': beginT, 'endT': endT, 'keyTPP': keyTPP };
				}
			}
			let stHi = JSON.stringify(chHi);

			let max = chData.arAnalys[dateD][hh].minmax.max;
			let min = chData.arAnalys[dateD][hh].minmax.min;


			let h0 = '0' + hh, h = h0.substr(-2, 2);

			chData.arExcel[ie] = { 'Date': dateP, 'Hour': h, 'Direction': '<49.99', 'Duration': beTimeBlockLow, 'MinMax': min }; ie++;
			chData.arExcel[ie] = { 'Date': dateP, 'Hour': h, 'Direction': '>50.01', 'Duration': beTimeBlockHi, 'MinMax': max }; ie++;
		}
	}

	setDefaltStyle(1, 0); // Візуалізація кнопки завантаження
}

//********************************************************************************
//БЛОК - 3 Заватаження Excel

//Excel з проаналізованою телеметрією ПДП
function downloadExcelFCR() {
	var flDate = chData.chDate.dateD;
	const worksheet = XLSX.utils.json_to_sheet(chData.arExcel);

	const workbook = {
		Sheets: { 'data': worksheet },
		SheetNames: ['data']
	};
	const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
	console.log(excelBuffer);
	saveAsExcel(excelBuffer, 'Activations_duration_at_BM_V1.02_' + flDate);
}

function saveAsExcel(buffer, filename) {
	const EXCEL_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
	const EXCEL_EXTENSION = ".xlsx";
	const data = new Blob([buffer], { type: EXCEL_TYPE });
	saveAs(data, filename + EXCEL_EXTENSION);
}