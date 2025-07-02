document.addEventListener('DOMContentLoaded', () => {
    const generateButton = document.getElementById('generate-button');
    const copyButton = document.getElementById('copy-button');
    const summarizeButton = document.getElementById('summarize-button');
    const bulletPointsButton = document.getElementById('bullet-points-button');
    const resultText = document.getElementById('result-text');

    generateButton.addEventListener('click', async () => {
        const subject = document.getElementById('mail-subject').value;
        const to = document.getElementById('mail-to').value;
        const purpose = document.getElementById('mail-purpose').value;
        const tone = document.getElementById('mail-tone').value;
        const manner = document.getElementById('mail-manner').value;

        const bodyPoints = Array.from(document.querySelectorAll('input[name="bodyPoints"]:checked'))
                                .map(checkbox => checkbox.value);

        const data = {
            subject,
            to,
            purpose,
            bodyPoints,
            tone,
            manner
        };

        try {
            const response = await fetch('/generate-email', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.message || 'メール生成に失敗しました。');
            }

            const result = await response.json();
            resultText.value = result.generatedEmail;
        } catch (error) {
            console.error('Error generating email:', error);
            resultText.value = `エラー: ${error.message}`;
        }
    });

    copyButton.addEventListener('click', () => {
        resultText.select();
        document.execCommand('copy');
        alert('生成結果をクリップボードにコピーしました！');
    });

    summarizeButton.addEventListener('click', async () => {
        const textToTransform = resultText.value;
        if (!textToTransform) {
            alert('生成結果がありません。');
            return;
        }

        try {
            const response = await fetch('/transform-text', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ text: textToTransform, type: 'summary' })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.message || '要約に失敗しました。');
            }

            const result = await response.json();
            resultText.value = result.transformedText;
        } catch (error) {
            console.error('Error summarizing text:', error);
            resultText.value = `エラー: ${error.message}`;
        }
    });

    bulletPointsButton.addEventListener('click', async () => {
        const textToTransform = resultText.value;
        if (!textToTransform) {
            alert('生成結果がありません。');
            return;
        }

        try {
            const response = await fetch('/transform-text', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ text: textToTransform, type: 'bullet_points' })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.message || '箇条書き変換に失敗しました。');
            }

            const result = await response.json();
            resultText.value = result.transformedText;
        } catch (error) {
            console.error('Error converting to bullet points:', error);
            resultText.value = `エラー: ${error.message}`;
        }
    });
});
