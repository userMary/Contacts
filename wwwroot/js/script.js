function saveContacts(format) {
    fetch(`/Contacts/SaveContacts?format=${format}`, { method: 'POST' })
        .then(response => response.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `contacts_data.${format}`;
            document.body.appendChild(a);
            a.click();
            a.remove();
        })
        .catch(error => console.error('Ошибка при сохранении:', error));
}