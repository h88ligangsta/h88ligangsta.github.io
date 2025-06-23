(function () {
    "use strict";

    // Данные адресов
    const data = [
        { organization: "Организация 1", address: "Улица 1, дом 1" },
        { organization: "ГКУ ДКР", address: "10-й мкр. Ясенева (Юго-Западный округ/Ясенево)" },
        { organization: "ГКУ ДКР", address: "10-й проезд Марьиной Рощи" },
        { organization: "ГКУ ДКР", address: "Попутная улица" },
        { organization: "ГКУ УКРиС", address: "Тупик (от Новорязанской ул. д.29 до парка им.Малютина)" },
        { organization: "АНО РГТ", address: "Ореховый бульвар, дом 37, корпус 3" },
        { organization: "ГКУ ДКР", address: "улица Воронцовские Пруды, дом 2, строение 1" },
        { organization: "ГКУ УКРиС", address: "Западный округ Тропарево-Никулино, без точного адреса" },
    ];

    Office.onReady(function (info) {
        if (info.host === Office.HostType.Outlook) {
            const input = document.getElementById("inputAddress");
            const suggestionList = document.getElementById("suggestionList");

            // Функция для вставки адреса в письмо
            function insertAddress(address) {
                return new Promise((resolve, reject) => {
                    // Проверяем доступность API для вставки текста
                    if (Office.context.mailbox.item.body.setSelectedDataAsync) {
                        Office.context.mailbox.item.body.setSelectedDataAsync(
                            address,
                            { coercionType: Office.CoercionType.Text },
                            (asyncResult) => {
                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                    resolve();
                                } else {
                                    reject(asyncResult.error);
                                }
                            }
                        );
                    } else {
                        // Альтернативный метод для некоторых версий Outlook
                        Office.context.mailbox.item.body.setSignatureAsync(
                            address,
                            { coercionType: Office.CoercionType.Html },
                            (asyncResult) => {
                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                                    resolve();
                                } else {
                                    reject(asyncResult.error);
                                }
                            }
                        );
                    }
                });
            }

            // Обработчик ввода
            input.addEventListener("input", function () {
                const query = input.value.trim().toLowerCase();
                suggestionList.innerHTML = "";

                if (query.length === 0) return;

                // Фильтрация данных
                const results = data.filter(item => 
                    item.address.toLowerCase().includes(query) || 
                    item.organization.toLowerCase().includes(query)
                );

                // Отображение результатов
                if (results.length > 0) {
                    results.forEach(item => {
                        const li = document.createElement("li");
                        li.innerHTML = `
                            <div class="address">${item.address}</div>
                            <div class="organization">${item.organization}</div>
                        `;
                        li.addEventListener("click", async () => {
                            try {
                                await insertAddress(item.address);
                                // Закрываем панель после успешной вставки
                                Office.addin.hide();
                            } catch (error) {
                                console.error("Ошибка при вставке адреса:", error);
                                suggestionList.innerHTML = `
                                    <li class="no-results">
                                        Ошибка при вставке адреса. Попробуйте вручную.
                                    </li>
                                `;
                            }
                        });
                        suggestionList.appendChild(li);
                    });
                } else {
                    suggestionList.innerHTML = '<li class="no-results">Адресов не найдено</li>';
                }
            });

            // Фокус на поле ввода при загрузке
            input.focus();

            // Обработчик нажатия клавиши Enter в поле ввода
            input.addEventListener("keydown", function (event) {
                if (event.key === "Enter") {
                    const firstSuggestion = suggestionList.querySelector("li");
                    if (firstSuggestion && !firstSuggestion.classList.contains("no-results")) {
                        firstSuggestion.click();
                    }
                }
            });
        }
    });

    // Функция для проверки, открыто ли письмо в режиме редактирования
    function isInComposeMode() {
        return Office.context.mailbox.item !== null && 
               (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) &&
               Office.context.mailbox.item.displayReplyForm !== undefined;
    }
})();