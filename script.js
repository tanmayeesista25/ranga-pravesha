const targetDate = new Date("July 18, 2026 15:00:00");

function updateCountdown() {

    const now = new Date();

    const diff = targetDate - now;

    if(diff <= 0){

        document.getElementById("countdown").innerHTML =
        "<div class='time-value'>Event Started!</div>";

        return;
    }

    const days =
        Math.floor(diff / (1000 * 60 * 60 * 24));

    const hours =
        Math.floor((diff % (1000 * 60 * 60 * 24))
        / (1000 * 60 * 60));

    const minutes =
        Math.floor((diff % (1000 * 60 * 60))
        / (1000 * 60));

    const seconds =
        Math.floor((diff % (1000 * 60))
        / 1000);

    document.getElementById("countdown").innerHTML = `

        <div class="time-block">
            <div class="time-value">${days}</div>
            <div class="time-label">Days</div>
        </div>

        <div class="time-separator">:</div>

        <div class="time-block">
            <div class="time-value">${hours}</div>
            <div class="time-label">Hours</div>
        </div>

        <div class="time-separator">:</div>

        <div class="time-block">
            <div class="time-value">${minutes}</div>
            <div class="time-label">Min</div>
        </div>

        <div class="time-separator">:</div>

        <div class="time-block">
            <div class="time-value">${seconds}</div>
            <div class="time-label">Sec</div>
        </div>
    `;
}

updateCountdown();

setInterval(updateCountdown,1000);

const galleryImages = [
    "images/DSC_3195.jpg",
    "images/DSC_3315.jpg",
    "images/DSC_3766.jpg",
    "images/DSC_3810.jpg",
    "images/DSC_3836.jpg",
    "images/DSC_3853.jpg",
    "images/DSC_4457.png",
    "images/DSC_4467.jpg",
    "images/DSC_4780.jpg",
    "images/DSC_4902.jpg",
    "images/DSC_4905.jpg",
    "images/DSC_4912.jpg",
    "images/DSC_4355.png",
];

let currentGalleryImage = 0;

window.addEventListener("load", () => {

    const image = document.getElementById("galleryImage");

setInterval(() => {

    image.style.opacity = 0;

    setTimeout(() => {

        currentGalleryImage =
            (currentGalleryImage + 1) %
            galleryImages.length;

        image.src =
            galleryImages[currentGalleryImage];

        image.style.opacity = 1;

    }, 800);

}, 5000);
});


document
.getElementById("rsvpForm")
.addEventListener("submit", function(e){

    e.preventDefault();

    const formData = new FormData();

formData.append(
    "entry.1191868580",
    document.getElementById("name").value
);

formData.append(
    "entry.1819345462",
    document.getElementById("email").value
);

formData.append(
    "entry.25566787",
    document.getElementById("phone").value
);

formData.append(
    "entry.1075416591",
    document.getElementById("attendance").value
);

formData.append(
    "entry.1301912319",
    document.getElementById("peopleCount").value
);

formData.append(
    "entry.124294925",
    document.getElementById("accessibility").value
);

formData.append(
    "entry.1257208760",
    document.getElementById("message").value
)

    fetch(
        "https://docs.google.com/forms/d/e/1FAIpQLSd3oiAoQOrBLI9mBcajq2e8F3dRdnVVS95NeNWAXtoggg-lQQ/formResponse",
        {
            method:"POST",
            mode:"no-cors",
            body:formData
        }
    )
    .then(() => {

    document
        .getElementById("rsvpForm")
        .style.display = "none";

    document
        .getElementById("successMessage")
        .style.display = "block";

    document
        .getElementById("successMessage")
        .scrollIntoView({
            behavior:"smooth",
            block:"center"
        });

    })
});
