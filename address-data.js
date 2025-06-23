// address-data.js
const addressData = [
  { organization: "Организация 1", address: "Улица 1, дом 1" },
  { organization: "ГКУ ДКР", address: "10-й мкр. Ясенева (Юго-Западный округ/Ясенево)" },
  { organization: "ГКУ ДКР", address: "10-й проезд Марьиной Рощи" },
  { organization: "ГКУ ДКР", address: "Попутная улица" },
  { organization: "ГКУ УКРиС", address: "Тупик (от Новорязанской ул. д.29 до парка им.Малютина)" },
  { organization: "АНО РГТ", address: "Ореховый бульвар, дом 37, корпус 3" },
  { organization: "ГКУ ДКР", address: "улица Воронцовские Пруды, дом 2, строение 1" },
  { organization: "ГКУ УКРиС", address: "Западный округ Тропарево-Никулино, без точного адреса" }
];

// Для использования в других файлах
if (typeof module !== 'undefined' && module.exports) {
  module.exports = addressData; // Для Node.js/тестов
} else {
  window.addressData = addressData; // Для браузера
}