const inputStudentName = document.getElementById("input-student-name")
const inputStudentGroup = document.getElementById("input-student-group")
const inputStudentId = document.getElementById("input-student-id")
const inputSubjectTournament = document.getElementById("input-subject-tournament")
const inputSubjectTeacher = document.getElementById("input-subject-teacher")
const inputGameOpponent = document.getElementById("input-game-opponent")
const inputGameWhite = document.getElementById("input-game-white")
const inputGameBlack = document.getElementById("input-game-black")

function handleFormSubmission(e) {
    e.preventDefault();

    const apiURL = new URL("/api/v1/report", window.location.origin)
    const requestBody = {
        student: {
            name: inputStudentName.value,
            group: inputStudentGroup.value,
            id: inputStudentId.value,
        },
        subject: {
            tournament: inputSubjectTournament.value,
            teacher: inputSubjectTeacher.value,
        },
        game: {
            opponent: inputGameOpponent.value,
            white_url: inputGameWhite.value,
            black_url: inputGameBlack.value,
        },
    }

    fetch(apiURL, {
        method: "POST",
        headers: {
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
    })
        .then((response) => {
            if (!response.ok) {
                return response.json().then((err) => Promise.reject(err));
            }
            return response.blob().then((blob) => ({
                blob: blob,
                filename: getFilenameFromHeader(response.headers.get("Content-Disposition"))
            }));
        })
        .then(({ blob, filename }) => {
            downloadBlob(blob, filename);
        })
        .catch(error => {
            console.error(error)
        });

}

function getFilenameFromHeader(header) {
    if (!header) return "report.xlsx";
    const match = header.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
    return match && match[1] ? decodeURIComponent(match[1].replace(/['"]/g, "")) : "report.xlsx";
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

document
    .getElementById("game-info-form")
    .addEventListener("submit", handleFormSubmission)
