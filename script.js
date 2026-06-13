
const targetDate=new Date('July 18, 2026 15:00:00');
function updateCountdown(){
 const diff=targetDate-new Date();
 if(diff<0) return;
 const days=Math.floor(diff/86400000);
 document.getElementById('countdown').innerHTML=days+' Days Until Ranga Pravesha';
}
updateCountdown();setInterval(updateCountdown,1000);

document.getElementById('rsvpForm').addEventListener('submit',e=>{
e.preventDefault();
document.getElementById('successMessage').style.display='block';
});
