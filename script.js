
const targetDate = new Date('July 18, 2026 15:00:00');
function updateCountdown(){
 const diff = targetDate - new Date();
 if(diff<0)return;
 const d=Math.floor(diff/86400000);
 document.getElementById('countdown').innerHTML=d+' days remaining';
}
updateCountdown();
setInterval(updateCountdown,1000);
