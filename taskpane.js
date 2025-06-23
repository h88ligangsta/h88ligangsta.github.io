(function () {
    "use strict";
    // Подключаем данные из отдельного файла
    const data = window.addressData;

    Office.onReady(function (info) {
        if (info.host === Office.HostType.Outlook) {
            const input = document.getElementById("inputAddress");
            const suggestionList = document.getElementById("suggestionList");

            // Остальной код остается без изменений...

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