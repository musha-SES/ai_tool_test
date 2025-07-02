document.getElementById('generateEmail').addEventListener('click', async () => {
    const subject = document.getElementById('subject').value;
    const to = document.getElementById('to').value;
    const purpose = document.getElementById('purpose').value;
    const bodyPoints = document.getElementById('bodyPoints').value.split(',').map(point => point.trim());
    const tone = document.getElementById('tone').value;
    const manner = document.getElementById('manner').value;

    try {
        const response = await fetch('http://localhost:3000/generate-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ subject, to, purpose, bodyPoints, tone, manner })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Failed to generate email');
        }

        const data = await response.json();
        document.getElementById('emailOutput').value = data.emailBody;
    } catch (error) {
        console.error('Error:', error);
        document.getElementById('emailOutput').value = `エラー: ${error.message}`;
    }
});

document.getElementById('transformButton').addEventListener('click', async () => {
    const text = document.getElementById('transformText').value;
    const type = document.getElementById('transformType').value;

    try {
        const response = await fetch('http://localhost:3000/transform-text', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ text, type })
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Failed to transform text');
        }

        const data = await response.json();
        document.getElementById('transformedOutput').value = data.transformedText;
    } catch (error) {
        console.error('Error:', error);
        document.getElementById('transformedOutput').value = `エラー: ${error.message}`;
    }
});
