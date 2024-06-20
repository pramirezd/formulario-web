document.getElementById('exampleForm').addEventListener('submit', async function(event) {
    event.preventDefault();

    const formData = {
        name: document.getElementById('name').value,
        age: document.getElementById('age').value,
        gender: document.getElementById('gender').value,
        hobbies: Array.from(document.querySelectorAll('input[name="hobbies"]:checked')).map(cb => cb.value)
    };

    const response = await fetch('/submit', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(formData)
    });

    const result = await response.json();
    alert(result.message);
});
