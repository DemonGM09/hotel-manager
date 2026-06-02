import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, CartesianGrid } from "recharts";

/* ═══════════════════ FONTS ═══════════════════ */
if (!document.getElementById("hf")) {
  const l = document.createElement("link"); l.id="hf"; l.rel="stylesheet";
  l.href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap";
  document.head.appendChild(l);
}

/* ═══════════════════ STORAGE ═══════════════════ */
const DB = {
  save: async (key, val) => { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} },
  load: async (key, def) => { try { const r = localStorage.getItem(key); return r ? JSON.parse(r) : def; } catch { return def; } },
};

/* ═══════════════════ CONSTANTS ═══════════════════ */
const FLOORS = { 1:["101","102"], 2:["201","202","203","204","205","206","207"], 3:["301","302","303","304","305","306","307"] };
const ALL_ROOMS = Object.values(FLOORS).flat();
const DRINKS = [
  {id:"nuoc", name:"Nước suối",icon:"💧",price:10000},{id:"bohuc",name:"Bò húc",icon:"🐂",price:10000},
  {id:"tra",  name:"Trà xanh", icon:"🍵",price:10000},{id:"c2",   name:"C2",    icon:"🧃",price:10000},
  {id:"sting",name:"Sting",    icon:"⚡",price:10000},{id:"bia",  name:"Bia",   icon:"🍺",price:10000},
];
const DEFAULT_STAFF = [
  {id:1,username:"admin",   password:"admin123",name:"Nguyễn Quản Lý",role:"admin"},
  {id:2,username:"nhanvien",password:"123456",  name:"Trần Nhân Viên",role:"staff"},
];
const STATUS_CFG = {
  available:{ label:"Trống",     color:"#4ade80",glow:"rgba(74,222,128,0.18)",  bg:"rgba(22,163,74,0.07)",  border:"rgba(74,222,128,0.25)",  dot:"#22c55e"},
  occupied: { label:"Có khách",  color:"#f87171",glow:"rgba(248,113,113,0.18)", bg:"rgba(220,38,38,0.07)",  border:"rgba(248,113,113,0.25)", dot:"#ef4444"},
  overnight:{ label:"Qua đêm",   color:"#a78bfa",glow:"rgba(167,139,250,0.18)", bg:"rgba(124,58,237,0.07)", border:"rgba(167,139,250,0.25)", dot:"#8b5cf6"},
  cleaning: { label:"Dọn phòng", color:"#fbbf24",glow:"rgba(251,191,36,0.18)",  bg:"rgba(217,119,6,0.07)",  border:"rgba(251,191,36,0.25)",  dot:"#f59e0b"},
  overtime: { label:"Quá giờ",   color:"#fb923c",glow:"rgba(251,146,60,0.18)",  bg:"rgba(234,88,12,0.07)",  border:"rgba(251,146,60,0.25)",  dot:"#f97316"},
};
const DAY_NAMES=["Chủ Nhật","Thứ Hai","Thứ Ba","Thứ Tư","Thứ Năm","Thứ Sáu","Thứ Bảy"];
const MONTH_NAMES=["Tháng 1","Tháng 2","Tháng 3","Tháng 4","Tháng 5","Tháng 6","Tháng 7","Tháng 8","Tháng 9","Tháng 10","Tháng 11","Tháng 12"];

/* ═══════════════════ HELPERS ═══════════════════ */
const isWeekend=(ts)=>[0,6].includes(new Date(ts).getDay());
const sameDay  =(a,b)=>new Date(a).toDateString()===new Date(b).toDateString();
const calcRoomBase=(room,now)=>{
  if(!room.checkInTime)return 0;
  const h=(now-room.checkInTime)/3600000;
  if(room.rateType==="hourly")return 120000+Math.max(0,Math.ceil(h)-2)*20000;
  return(isWeekend(room.checkInTime)?250000:220000)+Math.max(0,Math.ceil(h)-12)*20000;
};
const calcExtras=(e)=>(e||[]).reduce((s,i)=>s+i.qty*i.price,0);
const calcBill=(room,now)=>calcRoomBase(room,now)+calcExtras(room.extras);
const baseHours=(room)=>room.rateType==="hourly"?2:12;
const isOvertime=(room,now)=>room.checkInTime&&((now-room.checkInTime)/3600000)>baseHours(room);
const fmtMoney=(n)=>n.toLocaleString("vi-VN",{style:"currency",currency:"VND"});
const fmtDur=(ms)=>{if(!ms||ms<0)return"00:00:00";const h=Math.floor(ms/3600000),m=Math.floor((ms%3600000)/60000),s=Math.floor((ms%60000)/1000);return`${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`;};
const fmtDurS=(ms)=>{if(!ms||ms<0)return"0p";const h=Math.floor(ms/3600000),m=Math.floor((ms%3600000)/60000);return h>0?`${h}g ${m}p`:`${m}p`;};
const fmtTime=(ts)=>new Date(ts).toLocaleTimeString("vi-VN",{hour:"2-digit",minute:"2-digit",hour12:false});
const fmtDate=(ts)=>new Date(ts).toLocaleDateString("vi-VN",{day:"2-digit",month:"2-digit",year:"numeric"});
const fmtDob=(s)=>{if(!s)return"—";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;};
const EMPTY_ROOM=(n)=>({number:n,status:"available",checkInTime:null,rateType:null,guestName:"",cccd:"",dob:"",extras:[]});
const initRooms=()=>ALL_ROOMS.reduce((a,n)=>({...a,[n]:EMPTY_ROOM(n)}),{});
const EMPTY_FORM={name:"",cccd:"",dob:"",rate:"hourly"};

/* ═══════════════════ jsQR LOADER ═══════════════════ */
const loadJsQR=()=>new Promise((res,rej)=>{
  if(window.jsQR){res(window.jsQR);return;}
  const s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/jsqr/1.4.0/jsQR.min.js";
  s.onload=()=>res(window.jsQR); s.onerror=()=>rej(new Error("Không tải được jsQR"));
  document.head.appendChild(s);
});
const parseCCCDQR=(raw)=>{
  const p=raw.split("|");
  if(p.length<5)return null;
  const cccd=p[0]?.trim(),name=p[2]?.trim(),dobRaw=p[3]?.trim();
  if(!cccd||!/^\d{9,12}$/.test(cccd))return null;
  let dob="";
  if(dobRaw?.length===8){const d=dobRaw.slice(0,2),m=dobRaw.slice(2,4),y=dobRaw.slice(4,8);dob=`${y}-${m}-${d}`;}
  return{cccd,name,dob};
};

/* ═══════════════════ PRINT INVOICE ═══════════════════ */
const printInvoice=(h)=>{
  const extras=h.extras||[];
  const extrasRows=extras.map(e=>`<tr><td>${e.icon} ${e.name} ×${e.qty}</td><td style="text-align:right">${fmtMoney(e.qty*e.price)}</td></tr>`).join("");
  const html=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Hóa đơn ${h.room}</title>
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Segoe UI',Arial,sans-serif;max-width:320px;margin:20px auto;padding:20px;color:#1a1a2e;font-size:13px}
    .logo{text-align:center;margin-bottom:16px}
    .logo h1{font-size:20px;font-weight:700;letter-spacing:2px;color:#8c6a28}
    .logo p{font-size:11px;color:#666}
    .divider{border:none;border-top:1px dashed #ccc;margin:12px 0}
    .room-badge{text-align:center;font-size:32px;font-weight:700;color:#8c6a28;margin:8px 0}
    table{width:100%;border-collapse:collapse}
    td{padding:4px 0;vertical-align:top}
    td:first-child{color:#555}
    td:last-child{text-align:right;font-weight:500}
    .section-title{font-size:10px;font-weight:700;color:#888;letter-spacing:1px;margin:10px 0 4px}
    .total-row td{font-size:16px;font-weight:700;color:#8c6a28;padding-top:8px}
    .footer{text-align:center;font-size:11px;color:#888;margin-top:16px;line-height:1.6}
    @media print{body{margin:0}}
  </style></head><body>
  <div class="logo"><h1>◆ KHÁCH SẠN</h1><p>Cảm ơn quý khách đã sử dụng dịch vụ</p></div>
  <hr class="divider">
  <div class="room-badge">PHÒNG ${h.room}</div>
  <hr class="divider">
  <div class="section-title">THÔNG TIN KHÁCH</div>
  <table>
    <tr><td>Khách</td><td>${h.guestName||"—"}</td></tr>
    <tr><td>CCCD</td><td style="font-family:monospace">${h.cccd?h.cccd.replace(/(\d{4})(\d{4})(\d{4})/,"$1 $2 $3"):"—"}</td></tr>
    <tr><td>Ngày sinh</td><td>${fmtDob(h.dob)}</td></tr>
  </table>
  <hr class="divider">
  <div class="section-title">THỜI GIAN LƯU TRÚ</div>
  <table>
    <tr><td>Nhận phòng</td><td>${fmtTime(h.checkInTime)} ${fmtDate(h.checkInTime)}</td></tr>
    <tr><td>Trả phòng</td><td>${fmtTime(h.checkOutTime)} ${fmtDate(h.checkOutTime)}</td></tr>
    <tr><td>Thời gian</td><td>${fmtDurS(h.duration)}</td></tr>
    <tr><td>Loại</td><td>${h.rateType==="hourly"?"Ngắn giờ":"Theo ngày"}</td></tr>
  </table>
  <hr class="divider">
  <div class="section-title">CHI TIẾT HÓA ĐƠN</div>
  <table>
    <tr><td>Tiền phòng</td><td>${fmtMoney(h.roomBase)}</td></tr>
    ${extrasRows}
    <tr class="total-row"><td>TỔNG CỘNG</td><td>${fmtMoney(h.bill)}</td></tr>
  </table>
  <hr class="divider">
  <div class="footer">
    <p>Mã hóa đơn: #${String(h.id).slice(-6)}</p>
    <p>${fmtDate(h.checkOutTime)}</p>
    <p style="margin-top:8px">★ Hẹn gặp lại quý khách ★</p>
  </div>
  <script>window.onload=()=>{window.print();setTimeout(()=>window.close(),1000)}<\/script>
  </body></html>`;
  const w=window.open("","_blank","width=400,height=700");
  if(w){w.document.write(html);w.document.close();}
};

/* ═══════════════════ QR SCANNER ═══════════════════ */
function QRScanner({onResult,onClose}){
  /* mode: "pick" (initial) | "camera" | "file" */
  const [mode,setMode]=useState("pick");
  const [status,setStatus]=useState("idle"); // idle|loading|scanning|processing|success|error
  const [err,setErr]=useState("");
  const [previewSrc,setPreviewSrc]=useState(null); // for file mode
  const [fileResult,setFileResult]=useState(null); // decoded text from file
  const vidRef=useRef(),canRef=useRef(),streamRef=useRef(),timerRef=useRef(),fileRef=useRef();

  const stopCamera=useCallback(()=>{
    clearInterval(timerRef.current);
    streamRef.current?.getTracks().forEach(t=>t.stop());
    streamRef.current=null;
  },[]);

  /* ── start camera ── */
  useEffect(()=>{
    if(mode!=="camera")return;
    let alive=true;
    setStatus("loading");setErr("");
    (async()=>{
      try{
        const jsQR=await loadJsQR();
        const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:"environment",width:{ideal:1280}}});
        if(!alive){stream.getTracks().forEach(t=>t.stop());return;}
        streamRef.current=stream;
        vidRef.current.srcObject=stream;
        await vidRef.current.play();
        setStatus("scanning");
        timerRef.current=setInterval(()=>{
          const v=vidRef.current,c=canRef.current;
          if(!v||!c||v.readyState<2)return;
          c.width=v.videoWidth;c.height=v.videoHeight;
          const ctx=c.getContext("2d",{willReadFrequently:true});
          ctx.drawImage(v,0,0);
          const img=ctx.getImageData(0,0,c.width,c.height);
          const code=jsQR(img.data,img.width,img.height);
          if(code?.data){
            clearInterval(timerRef.current);
            setStatus("success");
            setTimeout(()=>{stopCamera();onResult(code.data);},500);
          }
        },150);
      }catch(e){
        if(!alive)return;
        setErr(e.name==="NotAllowedError"?"Vui lòng cho phép truy cập camera trong trình duyệt":e.message||"Lỗi camera");
        setStatus("error");
      }
    })();
    return()=>{alive=false;stopCamera();};
  },[mode]);

  /* ── decode image file ── */
  const decodeImageFile=async(file)=>{
    if(!file)return;
    setStatus("processing");setErr("");setFileResult(null);
    // show preview
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      const src=ev.target.result;
      setPreviewSrc(src);
      try{
        const jsQR=await loadJsQR();
        const img=new Image();
        img.onload=()=>{
          const c=canRef.current;
          c.width=img.width;c.height=img.height;
          const ctx=c.getContext("2d",{willReadFrequently:true});
          ctx.drawImage(img,0,0);
          const imgData=ctx.getImageData(0,0,c.width,c.height);
          const code=jsQR(imgData.data,imgData.width,imgData.height,{inversionAttempts:"attemptBoth"});
          if(code?.data){
            setFileResult(code.data);
            setStatus("success");
          }else{
            setErr("Không tìm thấy mã QR trong ảnh. Hãy chọn ảnh rõ hơn hoặc chụp gần hơn.");
            setStatus("error");
          }
        };
        img.onerror=()=>{setErr("Không đọc được file ảnh.");setStatus("error");};
        img.src=src;
      }catch(e){setErr(e.message||"Lỗi xử lý ảnh");setStatus("error");}
    };
    reader.readAsDataURL(file);
  };

  const handleFileChange=(e)=>{
    const file=e.target.files?.[0];
    if(file)decodeImageFile(file);
    e.target.value=""; // reset so same file can be picked again
  };

  const reset=()=>{setMode("pick");setStatus("idle");setErr("");setPreviewSrc(null);setFileResult(null);stopCamera();};

  /* ── PICK MODE ── */
  if(mode==="pick")return(
    <div style={{borderRadius:14,overflow:"hidden",background:"rgba(10,11,22,0.95)",border:"1px solid rgba(201,164,84,0.2)",padding:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:17,fontWeight:700,color:"#c9a455"}}>Quét mã QR CCCD</span>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#94a3b8",cursor:"pointer",width:28,height:28,borderRadius:"50%",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}}>✕</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {/* Camera option */}
        <div onClick={()=>setMode("camera")} style={{padding:"18px 12px",borderRadius:12,cursor:"pointer",textAlign:"center",border:"1px solid rgba(201,164,84,0.3)",background:"rgba(201,164,84,0.06)",transition:"all .15s",display:"flex",flexDirection:"column",alignItems:"center",gap:8,userSelect:"none"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(201,164,84,0.12)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(201,164,84,0.06)"}>
          <div style={{fontSize:36}}>📷</div>
          <div style={{fontSize:13,fontWeight:700,color:"#c9a455"}}>Quét bằng Camera</div>
          <div style={{fontSize:11,color:"#64748b",lineHeight:1.4}}>Mở camera và đưa mặt sau CCCD vào khung quét</div>
        </div>
        {/* File option */}
        <div onClick={()=>fileRef.current?.click()} style={{padding:"18px 12px",borderRadius:12,cursor:"pointer",textAlign:"center",border:"1px solid rgba(96,165,250,0.3)",background:"rgba(96,165,250,0.06)",transition:"all .15s",display:"flex",flexDirection:"column",alignItems:"center",gap:8,userSelect:"none"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(96,165,250,0.12)"}
          onMouseLeave={e=>e.currentTarget.style.background="rgba(96,165,250,0.06)"}>
          <div style={{fontSize:36}}>🖼️</div>
          <div style={{fontSize:13,fontWeight:700,color:"#60a5fa"}}>Chọn ảnh từ thư mục</div>
          <div style={{fontSize:11,color:"#64748b",lineHeight:1.4}}>Chọn ảnh chụp mặt sau CCCD từ thiết bị</div>
          <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{setMode("file");handleFileChange(e);}}/>
        </div>
      </div>
      <div style={{marginTop:12,padding:"9px 12px",background:"rgba(255,255,255,0.03)",borderRadius:8,fontSize:11,color:"#475569",lineHeight:1.6,textAlign:"center"}}>
        💡 Chụp rõ mã QR ở mặt sau CCCD — mã QR nằm ở góc dưới phải
      </div>
    </div>
  );

  /* ── FILE MODE ── */
  if(mode==="file")return(
    <div style={{borderRadius:14,overflow:"hidden",background:"rgba(10,11,22,0.95)",border:"1px solid rgba(96,165,250,0.2)",padding:16}}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
        <button onClick={reset} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#94a3b8",cursor:"pointer",padding:"4px 10px",borderRadius:8,fontSize:12,fontFamily:"inherit"}}>← Quay lại</button>
        <span style={{fontSize:13,fontWeight:600,color:"#60a5fa"}}>🖼️ Đọc QR từ ảnh</span>
      </div>
      <canvas ref={canRef} style={{display:"none"}}/>

      {status==="processing"&&(
        <div style={{padding:"32px 0",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",gap:12}}>
          <div style={{width:36,height:36,border:"3px solid rgba(96,165,250,0.3)",borderTop:"3px solid #60a5fa",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
          <span style={{color:"#60a5fa",fontSize:13}}>Đang phân tích ảnh…</span>
        </div>
      )}

      {previewSrc&&status!=="processing"&&(
        <div style={{marginBottom:12,borderRadius:10,overflow:"hidden",border:`2px solid ${status==="success"?"rgba(74,222,128,0.5)":"rgba(248,113,113,0.4)"}`}}>
          <img src={previewSrc} alt="preview" style={{width:"100%",maxHeight:220,objectFit:"contain",background:"#000",display:"block"}}/>
        </div>
      )}

      {status==="success"&&fileResult&&(
        <div style={{marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px",background:"rgba(74,222,128,0.1)",border:"1px solid rgba(74,222,128,0.25)",borderRadius:10,marginBottom:10}}>
            <span style={{fontSize:20}}>✅</span>
            <span style={{fontSize:13,color:"#4ade80",fontWeight:600}}>Đọc QR thành công!</span>
          </div>
          <button onClick={()=>onResult(fileResult)} style={{width:"100%",padding:"12px",borderRadius:10,background:"linear-gradient(135deg,#c9a455,#8c6a28)",border:"none",color:"#0a0b12",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif"}}>
            Dùng thông tin này
          </button>
        </div>
      )}

      {status==="error"&&(
        <div style={{marginBottom:12}}>
          <div style={{padding:"10px 12px",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.25)",borderRadius:10,marginBottom:10}}>
            <div style={{fontSize:13,color:"#f87171"}}>{err}</div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <button onClick={()=>fileRef.current?.click()} style={{padding:"10px",borderRadius:10,background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.25)",color:"#60a5fa",cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>
              🖼️ Chọn ảnh khác
            </button>
            <button onClick={()=>{reset();setTimeout(()=>setMode("camera"),50);}} style={{padding:"10px",borderRadius:10,background:"rgba(201,164,84,0.1)",border:"1px solid rgba(201,164,84,0.25)",color:"#c9a455",cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>
              📷 Dùng camera
            </button>
          </div>
          <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleFileChange}/>
        </div>
      )}

      {status==="idle"&&(
        <div style={{padding:"28px 0",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",gap:10}}>
          <div style={{fontSize:40}}>🖼️</div>
          <div style={{fontSize:13,color:"#64748b"}}>Đang chờ chọn ảnh…</div>
          <input ref={fileRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleFileChange}/>
          <button onClick={()=>fileRef.current?.click()} style={{padding:"9px 20px",borderRadius:10,background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.25)",color:"#60a5fa",cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:600}}>
            Chọn ảnh từ thiết bị
          </button>
        </div>
      )}
    </div>
  );

  /* ── CAMERA MODE ── */
  return(
    <div style={{borderRadius:14,overflow:"hidden",background:"rgba(10,11,22,0.95)",border:"1px solid rgba(201,164,84,0.2)"}}>
      {/* Camera header */}
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"10px 14px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
        <button onClick={reset} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#94a3b8",cursor:"pointer",padding:"4px 10px",borderRadius:8,fontSize:12,fontFamily:"inherit"}}>← Quay lại</button>
        <span style={{fontSize:13,fontWeight:600,color:"#c9a455"}}>📷 Quét bằng Camera</span>
        <div style={{marginLeft:"auto",display:"flex",gap:6}}>
          {/* Switch to file */}
          <button onClick={()=>{stopCamera();setMode("file");setStatus("idle");}} style={{padding:"4px 10px",borderRadius:8,background:"rgba(96,165,250,0.08)",border:"1px solid rgba(96,165,250,0.2)",color:"#60a5fa",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>🖼️ Chọn ảnh</button>
        </div>
      </div>
      {/* Camera view */}
      <div style={{position:"relative",aspectRatio:"4/3",background:"#000"}}>
        <video ref={vidRef} playsInline muted style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}/>
        <canvas ref={canRef} style={{display:"none"}}/>
        {status==="success"&&<div style={{position:"absolute",inset:0,background:"rgba(74,222,128,0.45)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:52}}>✅</div>}
        {status==="scanning"&&(
          <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",pointerEvents:"none"}}>
            <div style={{width:190,height:190,position:"relative"}}>
              {[0,1,2,3].map(i=>(
                <div key={i} style={{position:"absolute",width:26,height:26,
                  top:i<2?"0":"auto",bottom:i>=2?"0":"auto",
                  left:i%2===0?"0":"auto",right:i%2===1?"0":"auto",
                  borderTop:i<2?"3px solid #c9a455":"none",borderBottom:i>=2?"3px solid #c9a455":"none",
                  borderLeft:i%2===0?"3px solid #c9a455":"none",borderRight:i%2===1?"3px solid #c9a455":"none"}}/>
              ))}
              {/* animated scan line */}
              <div style={{position:"absolute",left:4,right:4,height:2,background:"linear-gradient(90deg,transparent,#c9a455,transparent)",animation:"scanline 1.6s ease-in-out infinite"}}/>
            </div>
            <div style={{position:"absolute",bottom:14,fontSize:11,color:"#c9a455",background:"rgba(0,0,0,0.65)",padding:"5px 14px",borderRadius:20,backdropFilter:"blur(4px)"}}>
              Đưa mã QR mặt sau CCCD vào khung
            </div>
          </div>
        )}
        {status==="loading"&&(
          <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.75)",gap:12}}>
            <div style={{width:34,height:34,border:"3px solid rgba(201,164,84,0.3)",borderTop:"3px solid #c9a455",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
            <span style={{color:"#c9a455",fontSize:12}}>Khởi động camera…</span>
          </div>
        )}
        {status==="error"&&(
          <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.88)",gap:12,padding:20,textAlign:"center"}}>
            <div style={{fontSize:34}}>📷</div>
            <div style={{color:"#f87171",fontSize:13,lineHeight:1.5}}>{err}</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>{stopCamera();setMode("file");setStatus("idle");}} style={{padding:"7px 14px",borderRadius:8,background:"rgba(96,165,250,0.12)",border:"1px solid rgba(96,165,250,0.25)",color:"#60a5fa",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>🖼️ Chọn ảnh</button>
              <button onClick={onClose} style={{padding:"7px 14px",borderRadius:8,background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.15)",color:"#e2e8f0",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Đóng</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════ TOAST SYSTEM ═══════════════════ */
function ToastContainer({toasts}){
  return(<div style={{position:"fixed",bottom:20,right:20,zIndex:999,display:"flex",flexDirection:"column",gap:8}}>
    {toasts.map(t=>(
      <div key={t.id} style={{padding:"10px 16px",borderRadius:10,background:t.type==="success"?"rgba(22,163,74,0.9)":t.type==="warn"?"rgba(234,88,12,0.9)":"rgba(30,32,50,0.95)",border:`1px solid ${t.type==="success"?"rgba(74,222,128,0.4)":t.type==="warn"?"rgba(251,146,60,0.4)":"rgba(255,255,255,0.1)"}`,color:"#fff",fontSize:13,fontFamily:"'DM Sans',sans-serif",backdropFilter:"blur(8px)",boxShadow:"0 4px 16px rgba(0,0,0,0.3)",animation:"slideUp .3s ease",maxWidth:280}}>
        {t.msg}
      </div>
    ))}
  </div>);
}

/* ═══════════════════ LOGIN SCREEN ═══════════════════ */
function LoginScreen({onLogin,theme}){
  const [u,setU]=useState(""),  [p,setP]=useState(""), [err,setErr]=useState(""), [loading,setLoading]=useState(false);
  const t=theme==="dark";
  const login=async()=>{
    setLoading(true);setErr("");
    const staff=await DB.load("staff",DEFAULT_STAFF);
    const found=staff.find(s=>s.username===u.trim()&&s.password===p);
    setTimeout(()=>{
      if(found){onLogin(found);}
      else{setErr("Sai tên đăng nhập hoặc mật khẩu");setLoading(false);}
    },400);
  };
  return(
    <div style={{minHeight:"100vh",background:t?"#0b0c16":"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{width:"100%",maxWidth:360,background:t?"#10111d":"#fff",border:`1px solid ${t?"rgba(201,164,84,0.2)":"rgba(0,0,0,0.08)"}`,borderRadius:20,padding:36,boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:32,fontWeight:700,color:"#c9a455",letterSpacing:3}}>◆ HOTEL</div>
          <div style={{fontSize:13,color:t?"#64748b":"#94a3b8",marginTop:6}}>Hệ thống quản lý khách sạn</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div>
            <label style={{fontSize:11,fontWeight:600,color:t?"#64748b":"#94a3b8",letterSpacing:.5,display:"block",marginBottom:6}}>TÊN ĐĂNG NHẬP</label>
            <input value={u} onChange={e=>setU(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} placeholder="admin / nhanvien"
              style={{width:"100%",padding:"11px 14px",background:t?"rgba(255,255,255,0.04)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:10,color:t?"#e2e8f0":"#1e293b",fontSize:14,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          <div>
            <label style={{fontSize:11,fontWeight:600,color:t?"#64748b":"#94a3b8",letterSpacing:.5,display:"block",marginBottom:6}}>MẬT KHẨU</label>
            <input type="password" value={p} onChange={e=>setP(e.target.value)} onKeyDown={e=>e.key==="Enter"&&login()} placeholder="••••••••"
              style={{width:"100%",padding:"11px 14px",background:t?"rgba(255,255,255,0.04)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:10,color:t?"#e2e8f0":"#1e293b",fontSize:14,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          {err&&<div style={{fontSize:12,color:"#f87171",textAlign:"center"}}>{err}</div>}
          <button onClick={login} disabled={loading||!u||!p}
            style={{padding:"12px",borderRadius:10,background:"linear-gradient(135deg,#c9a455,#8c6a28)",border:"none",color:"#0a0b12",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",opacity:loading||!u||!p?0.6:1,marginTop:4}}>
            {loading?"Đang đăng nhập…":"Đăng nhập"}
          </button>
        </div>
        <div style={{marginTop:24,padding:14,background:t?"rgba(255,255,255,0.02)":"#f8fafc",borderRadius:10,border:`1px solid ${t?"rgba(255,255,255,0.06)":"#e2e8f0"}`}}>
          <div style={{fontSize:11,color:t?"#475569":"#94a3b8",marginBottom:6,fontWeight:600}}>TÀI KHOẢN MẶC ĐỊNH</div>
          <div style={{fontSize:12,color:t?"#64748b":"#94a3b8",lineHeight:1.8}}>
            👑 Quản lý: <code style={{color:"#c9a455"}}>admin / admin123</code><br/>
            👤 Nhân viên: <code style={{color:"#94a3b8"}}>nhanvien / 123456</code>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════ STATS VIEW ═══════════════════ */
function StatsView({history,theme}){
  const t=theme==="dark";
  const now=new Date();
  const [selMonth,setSelMonth]=useState(now.getMonth());
  const [selYear, setSelYear] =useState(now.getFullYear());

  const monthData=useMemo(()=>{
    const filtered=history.filter(h=>{
      const d=new Date(h.checkOutTime);
      return d.getMonth()===selMonth&&d.getFullYear()===selYear;
    });
    const byDay={};
    filtered.forEach(h=>{
      const day=new Date(h.checkOutTime).getDate();
      byDay[day]=(byDay[day]||0)+h.bill;
    });
    const daysInMonth=new Date(selYear,selMonth+1,0).getDate();
    return Array.from({length:daysInMonth},(_,i)=>({day:`${i+1}`,rev:byDay[i+1]||0}));
  },[history,selMonth,selYear]);

  const summary=useMemo(()=>{
    const filtered=history.filter(h=>{const d=new Date(h.checkOutTime);return d.getMonth()===selMonth&&d.getFullYear()===selYear;});
    const total=filtered.reduce((s,h)=>s+h.bill,0);
    const avgDur=filtered.length?filtered.reduce((s,h)=>s+h.duration,0)/filtered.length:0;
    const roomCount={};
    filtered.forEach(h=>{roomCount[h.room]=(roomCount[h.room]||0)+1;});
    const topRoom=Object.entries(roomCount).sort((a,b)=>b[1]-a[1])[0];
    return{total,count:filtered.length,avgDur,topRoom:topRoom?topRoom[0]:"—",extrasTotal:filtered.reduce((s,h)=>s+(h.extrasTotal||0),0)};
  },[history,selMonth,selYear]);

  const card=(label,value,sub,clr="#c9a455")=>(
    <div style={{padding:"16px 18px",background:t?"rgba(255,255,255,0.03)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.08)":"#e2e8f0"}`,borderRadius:12}}>
      <div style={{fontSize:10,color:t?"#475569":"#94a3b8",fontWeight:600,letterSpacing:.5,marginBottom:6}}>{label}</div>
      <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:700,color:clr}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:t?"#64748b":"#94a3b8",marginTop:3}}>{sub}</div>}
    </div>
  );

  return(
    <div style={{padding:"24px 20px"}}>
      {/* Month selector */}
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24,flexWrap:"wrap"}}>
        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:700,color:"#c9a455"}}>📊 Thống kê doanh thu</div>
        <div style={{display:"flex",gap:8,marginLeft:"auto"}}>
          <select value={selMonth} onChange={e=>setSelMonth(+e.target.value)}
            style={{padding:"7px 12px",background:t?"#1e1f2e":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:8,color:t?"#e2e8f0":"#1e293b",fontSize:13,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",outline:"none"}}>
            {MONTH_NAMES.map((n,i)=><option key={i} value={i}>{n}</option>)}
          </select>
          <select value={selYear} onChange={e=>setSelYear(+e.target.value)}
            style={{padding:"7px 12px",background:t?"#1e1f2e":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:8,color:t?"#e2e8f0":"#1e293b",fontSize:13,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",outline:"none"}}>
            {[2024,2025,2026].map(y=><option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </div>

      {/* Summary cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:24}}>
        {card("TỔNG DOANH THU",fmtMoney(summary.total),`${summary.count} lượt trả phòng`)}
        {card("ĐỒ UỐNG",fmtMoney(summary.extrasTotal),"Tổng phát sinh","#f59e0b")}
        {card("THỜI GIAN TB",fmtDurS(summary.avgDur),"Mỗi lượt lưu trú","#60a5fa")}
        {card("PHÒNG HOT",summary.topRoom,"Nhiều lượt nhất","#a78bfa")}
      </div>

      {/* Chart */}
      <div style={{background:t?"rgba(255,255,255,0.02)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:14,padding:"20px 16px",marginBottom:24}}>
        <div style={{fontSize:12,color:t?"#64748b":"#94a3b8",fontWeight:600,letterSpacing:.5,marginBottom:16}}>DOANH THU THEO NGÀY (đ)</div>
        {summary.count===0?(
          <div style={{height:200,display:"flex",alignItems:"center",justifyContent:"center",color:t?"#334155":"#94a3b8",fontSize:14}}>Chưa có dữ liệu tháng này</div>
        ):(
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={monthData} margin={{top:4,right:8,left:0,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke={t?"rgba(255,255,255,0.05)":"rgba(0,0,0,0.05)"}/>
              <XAxis dataKey="day" tick={{fontSize:10,fill:t?"#64748b":"#94a3b8"}} axisLine={false} tickLine={false}/>
              <YAxis tickFormatter={v=>v>=1000?`${v/1000}k`:v} tick={{fontSize:10,fill:t?"#64748b":"#94a3b8"}} axisLine={false} tickLine={false} width={40}/>
              <Tooltip formatter={v=>[fmtMoney(v),"Doanh thu"]} contentStyle={{background:t?"#1e1f2e":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:8,color:t?"#e2e8f0":"#1e293b",fontSize:12}}/>
              <Bar dataKey="rev" fill="#c9a455" radius={[4,4,0,0]} maxBarSize={24}/>
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════ SETTINGS VIEW ═══════════════════ */
function SettingsView({theme,setTheme,settings,setSettings,history,rooms,totalRevenue,onLogout,currentUser,addToast}){
  const t=theme==="dark";
  const fileRef=useRef();

  const exportJSON=()=>{
    const data={rooms,history,totalRevenue,exportedAt:new Date().toISOString(),version:"2.0"};
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`hotel-backup-${fmtDate(Date.now()).replace(/\//g,"-")}.json`;a.click();
    addToast("✅ Đã xuất file backup","success");
  };

  const importJSON=(e)=>{
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      try{
        const data=JSON.parse(ev.target.result);
        if(data.rooms)await DB.save("rooms",data.rooms);
        if(data.history)await DB.save("history",data.history);
        if(data.totalRevenue!=null)await DB.save("totalRevenue",data.totalRevenue);
        addToast("✅ Đã nhập backup thành công — Tải lại trang để áp dụng","success");
      }catch{addToast("❌ File backup không hợp lệ","error");}
    };
    reader.readAsText(file);
  };

  const exportExcel=()=>{
    const rows=history.map(h=>({
      "Phòng":h.room,"Khách":h.guestName||"","Số CCCD":h.cccd||"","Ngày sinh":fmtDob(h.dob),
      "Giờ vào":fmtTime(h.checkInTime)+" "+fmtDate(h.checkInTime),
      "Giờ ra":fmtTime(h.checkOutTime)+" "+fmtDate(h.checkOutTime),
      "Thời gian":fmtDurS(h.duration),"Loại":h.rateType==="hourly"?"Ngắn giờ":"Theo ngày",
      "Tiền phòng (đ)":h.roomBase,"Đồ uống (đ)":h.extrasTotal||0,"Tổng (đ)":h.bill,
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Lịch sử");
    XLSX.writeFile(wb,`hotel-history-${fmtDate(Date.now()).replace(/\//g,"-")}.xlsx`);
    addToast("✅ Đã xuất file Excel","success");
  };

  const Row=({label,children})=>(
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"14px 0",borderBottom:`1px solid ${t?"rgba(255,255,255,0.06)":"#f1f5f9"}`}}>
      <span style={{fontSize:14,color:t?"#cbd5e1":"#475569"}}>{label}</span>
      {children}
    </div>
  );
  const Toggle=({val,onChange})=>(
    <div onClick={()=>onChange(!val)} style={{width:44,height:24,borderRadius:12,background:val?"#c9a455":"rgba(255,255,255,0.1)",cursor:"pointer",position:"relative",transition:"background .2s",border:`1px solid ${val?"#a8834a":"rgba(255,255,255,0.15)"}`}}>
      <div style={{position:"absolute",top:3,left:val?21:3,width:16,height:16,borderRadius:"50%",background:"#fff",transition:"left .2s"}}/>
    </div>
  );

  return(
    <div style={{padding:"24px 20px"}}>
      <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:700,color:"#c9a455",marginBottom:24}}>⚙️ Cài đặt</div>

      {/* User info */}
      <div style={{padding:"16px 20px",background:t?"rgba(201,164,84,0.07)":"#fffbeb",border:`1px solid ${t?"rgba(201,164,84,0.2)":"rgba(201,164,84,0.3)"}`,borderRadius:12,marginBottom:20,display:"flex",alignItems:"center",gap:14}}>
        <div style={{width:44,height:44,borderRadius:"50%",background:"linear-gradient(135deg,#c9a455,#8c6a28)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>{currentUser.role==="admin"?"👑":"👤"}</div>
        <div>
          <div style={{fontWeight:600,color:t?"#e2e8f0":"#1e293b"}}>{currentUser.name}</div>
          <div style={{fontSize:12,color:t?"#64748b":"#94a3b8"}}>{currentUser.role==="admin"?"Quản lý":"Nhân viên"} · @{currentUser.username}</div>
        </div>
        <button onClick={onLogout} style={{marginLeft:"auto",padding:"7px 14px",borderRadius:8,background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.2)",color:"#f87171",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Đăng xuất</button>
      </div>

      {/* Giao diện */}
      <div style={{background:t?"rgba(255,255,255,0.02)":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:12,padding:"4px 20px",marginBottom:16}}>
        <div style={{fontSize:11,color:t?"#475569":"#94a3b8",fontWeight:600,letterSpacing:.5,padding:"12px 0 6px"}}>GIAO DIỆN</div>
        <Row label="🌙 Chế độ tối"><Toggle val={theme==="dark"} onChange={v=>setTheme(v?"dark":"light")}/></Row>
      </div>

      {/* Auto checkout */}
      <div style={{background:t?"rgba(255,255,255,0.02)":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:12,padding:"4px 20px",marginBottom:16}}>
        <div style={{fontSize:11,color:t?"#475569":"#94a3b8",fontWeight:600,letterSpacing:.5,padding:"12px 0 6px"}}>KIỂM SOÁT PHÒNG</div>
        <Row label="⏰ Cảnh báo quá giờ"><Toggle val={settings.overtimeAlert} onChange={v=>setSettings(s=>({...s,overtimeAlert:v}))}/></Row>
        <Row label="🤖 Tự động trả phòng khi quá giờ"><Toggle val={settings.autoCheckout} onChange={v=>setSettings(s=>({...s,autoCheckout:v}))}/></Row>
        <Row label={<span>⏱ Thời gian ân hạn sau khi quá giờ<br/><span style={{fontSize:11,color:t?"#475569":"#94a3b8"}}>Tự động trả phòng sau X phút</span></span>}>
          <select value={settings.graceMins||30} onChange={e=>setSettings(s=>({...s,graceMins:+e.target.value}))}
            style={{padding:"6px 10px",background:t?"#1e1f2e":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"}`,borderRadius:8,color:t?"#e2e8f0":"#1e293b",fontSize:13,fontFamily:"inherit",cursor:"pointer",outline:"none"}}>
            {[15,30,60].map(m=><option key={m} value={m}>{m} phút</option>)}
          </select>
        </Row>
      </div>

      {/* Data */}
      <div style={{background:t?"rgba(255,255,255,0.02)":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:12,padding:"4px 20px",marginBottom:16}}>
        <div style={{fontSize:11,color:t?"#475569":"#94a3b8",fontWeight:600,letterSpacing:.5,padding:"12px 0 6px"}}>DỮ LIỆU & SAO LƯU</div>
        <Row label="📊 Xuất lịch sử ra Excel">
          <button onClick={exportExcel} style={{padding:"7px 14px",borderRadius:8,background:"rgba(74,222,128,0.1)",border:"1px solid rgba(74,222,128,0.2)",color:"#4ade80",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Xuất Excel</button>
        </Row>
        <Row label="💾 Sao lưu dữ liệu (JSON)">
          <button onClick={exportJSON} style={{padding:"7px 14px",borderRadius:8,background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.2)",color:"#60a5fa",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Xuất Backup</button>
        </Row>
        <Row label="📂 Khôi phục từ file backup">
          <>
            <input ref={fileRef} type="file" accept=".json" onChange={importJSON} style={{display:"none"}}/>
            <button onClick={()=>fileRef.current.click()} style={{padding:"7px 14px",borderRadius:8,background:"rgba(167,139,250,0.1)",border:"1px solid rgba(167,139,250,0.2)",color:"#a78bfa",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>Nhập Backup</button>
          </>
        </Row>
      </div>

      <div style={{fontSize:11,color:t?"#334155":"#cbd5e1",textAlign:"center",marginTop:8}}>
        📡 Dữ liệu được lưu tự động trên cloud · {history.length} bản ghi · v2.0
      </div>
    </div>
  );
}

/* ═══════════════════ CSS ═══════════════════ */
const buildCSS=(t)=>`
  *{box-sizing:border-box;margin:0;padding:0;}
  :root{scrollbar-width:thin;scrollbar-color:${t?"#1e1f2e":"#cbd5e1"} transparent;}
  @keyframes pulse    {0%,100%{opacity:1;transform:scale(1)}50%{opacity:.6;transform:scale(.85)}}
  @keyframes fadeIn   {from{opacity:0}to{opacity:1}}
  @keyframes fadeSlide{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:translateY(0)}}
  @keyframes slideUp  {from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
  @keyframes countUp  {from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}
  @keyframes rowAppear{from{opacity:0;transform:translateX(-8px)}to{opacity:1;transform:translateX(0)}}
  @keyframes spin     {to{transform:rotate(360deg)}}
  @keyframes pop      {0%{transform:scale(1)}40%{transform:scale(1.2)}100%{transform:scale(1)}}
  @keyframes qrPulse  {0%,100%{box-shadow:0 0 0 0 rgba(201,164,84,0.4)}70%{box-shadow:0 0 0 8px rgba(201,164,84,0)}}
  @keyframes scanline {0%{top:10%}50%{top:85%}100%{top:10%}}
  .room-card{transition:transform .15s,box-shadow .15s;cursor:pointer;user-select:none;}
  .room-card:hover{transform:translateY(-3px);}
  .room-card:active{transform:translateY(0);}
  .pulse{animation:pulse 1.8s ease-in-out infinite;}
  .modal-bg{animation:fadeIn .2s ease;}
  .modal-box{animation:fadeSlide .2s ease;}
  .count-up{animation:countUp .4s ease;}
  .pop{animation:pop .25s ease;}
  .btn{width:100%;padding:13px;border-radius:10px;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:600;font-size:14px;border:none;transition:opacity .15s,transform .1s;}
  .btn:hover{opacity:.9;}.btn:active{transform:scale(.98);}
  .fl{font-size:11px;color:${t?"#64748b":"#94a3b8"};display:block;margin-bottom:6px;font-weight:600;letter-spacing:.5px;}
  .fi{font-family:'DM Sans',sans-serif;width:100%;padding:10px 14px;background:${t?"rgba(255,255,255,0.04)":"#f8fafc"};border:1px solid ${t?"rgba(255,255,255,0.1)":"#e2e8f0"};border-radius:8px;color:${t?"#e2e8f0":"#1e293b"};font-size:14px;outline:none;transition:border-color .15s;}
  .fi:focus{border-color:rgba(201,164,84,.5);}
  .fi::placeholder{color:${t?"#475569":"#94a3b8"};}
  input[type=date]::-webkit-calendar-picker-indicator{filter:${t?"invert(.4)":"none"};cursor:pointer;}
  .sp{display:flex;align-items:center;gap:7px;padding:5px 10px;border-radius:20px;font-size:12px;font-weight:500;}
  .mi{overflow-y:auto;max-height:calc(100vh - 60px);}
  .mi::-webkit-scrollbar{width:4px;}
  .mi::-webkit-scrollbar-thumb{background:${t?"#1e1f2e":"#e2e8f0"};border-radius:4px;}
  .ht{width:100%;border-collapse:collapse;}
  .ht th{font-size:10px;font-weight:600;color:${t?"#475569":"#94a3b8"};letter-spacing:.6px;padding:0 10px 10px;text-align:left;}
  .ht td{padding:10px;font-size:12px;color:${t?"#cbd5e1":"#475569"};border-top:1px solid ${t?"rgba(255,255,255,0.05)":"#f1f5f9"};vertical-align:middle;}
  .ht tr:hover td{background:${t?"rgba(255,255,255,0.02)":"#f8fafc"};}
  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;}
  .db{display:flex;flex-direction:column;align-items:center;gap:4px;padding:10px 8px;border-radius:10px;cursor:pointer;border:1px solid rgba(255,255,255,0.08);background:rgba(255,255,255,0.03);transition:all .15s;user-select:none;flex:1;min-width:60px;}
  .db:hover{border-color:rgba(201,164,84,.35);background:rgba(201,164,84,0.06);}
  .qb{display:flex;align-items:center;justify-content:center;gap:8px;width:100%;padding:11px;border-radius:10px;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:600;font-size:13px;border:1px dashed rgba(201,164,84,0.5);background:rgba(201,164,84,0.05);color:#c9a455;transition:all .15s;animation:qrPulse 2.5s infinite;}
  .qb:hover{background:rgba(201,164,84,0.1);border-style:solid;}
  .tb{padding:7px 14px;border-radius:8px;cursor:pointer;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;border:1px solid rgba(255,255,255,0.08);background:transparent;color:${t?"#64748b":"#94a3b8"};transition:all .15s;}
  .tb.act{background:rgba(201,164,84,0.12);border-color:rgba(201,164,84,.35);color:#c9a455;}
  .nav-btn{display:flex;flex-direction:column;align-items:center;gap:3px;padding:8px 16px;border-radius:10px;cursor:pointer;border:none;background:transparent;color:${t?"#64748b":"#94a3b8"};font-family:'DM Sans',sans-serif;font-size:11px;font-weight:500;transition:all .15s;flex:1;}
  .nav-btn.act{color:#c9a455;background:rgba(201,164,84,0.08);}
  .nav-btn:hover:not(.act){background:${t?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.04)"};}
  @media(max-width:640px){
    .room-grid{gap:8px!important;}
    .room-card{width:calc(50% - 4px)!important;min-width:0!important;}
    .header-row{flex-direction:column!important;gap:12px!important;}
    .rev-cards{gap:8px!important;}
    .rev-card{min-width:0!important;flex:1!important;}
    .stat-bar{flex-wrap:wrap!important;}
  }
`;

/* ═══════════════════ MAIN APP ═══════════════════ */
function HotelApp({currentUser,onLogout,theme,setTheme}){
  const [rooms,       setRooms]       = useState(initRooms);
  const [tick,        setTick]        = useState(Date.now());
  const [totalRev,    setTotalRev]    = useState(0);
  const [history,     setHistory]     = useState([]);
  const [settings,    setSettings]    = useState({overtimeAlert:true,autoCheckout:false,graceMins:30});
  const [view,        setView]        = useState("rooms"); // rooms|stats|settings
  const [sel,         setSel]         = useState(null);
  const [modal,       setModal]       = useState(null);   // checkin|detail
  const [form,        setForm]        = useState(EMPTY_FORM);
  const [showQR,      setShowQR]      = useState(false);
  const [qrScanned,   setQrScanned]   = useState(false);
  const [detailTab,   setDetailTab]   = useState("info");
  const [histOpen,    setHistOpen]    = useState(true);
  const [popId,       setPopId]       = useState(null);
  const [toasts,      setToasts]      = useState([]);
  const [loaded,      setLoaded]      = useState(false);
  const alertedRef    = useRef(new Set());

  const addToast=useCallback((msg,type="info")=>{
    const id=Date.now();
    setToasts(t=>[...t,{id,msg,type}]);
    setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),4000);
  },[]);

  /* Load from storage */
  useEffect(()=>{
    (async()=>{
      const [r,h,rev,sett]=await Promise.all([
        DB.load("rooms",null),DB.load("history",[]),
        DB.load("totalRevenue",0),DB.load("settings",settings),
      ]);
      if(r){
        const merged=initRooms();
        Object.keys(merged).forEach(k=>{if(r[k])merged[k]={...merged[k],...r[k],extras:r[k].extras||[]};});
        setRooms(merged);
      }
      setHistory(h);setTotalRev(rev);setSettings(sett);setLoaded(true);
    })();
  },[]);

  /* Auto-save rooms */
  useEffect(()=>{ if(loaded)DB.save("rooms",rooms); },[rooms,loaded]);
  useEffect(()=>{ if(loaded)DB.save("history",history); },[history,loaded]);
  useEffect(()=>{ if(loaded)DB.save("totalRevenue",totalRev); },[totalRev,loaded]);
  useEffect(()=>{ if(loaded)DB.save("settings",settings); },[settings,loaded]);

  /* Clock + midnight + auto-checkout */
  useEffect(()=>{
    const id=setInterval(()=>{
      const now=Date.now(); setTick(now);
      setRooms(prev=>{
        let changed=false; const next={...prev};
        for(const r of Object.values(next)){
          if((r.status==="occupied"||r.status==="overtime")&&r.checkInTime){
            // midnight → overnight
            if(r.status==="occupied"&&new Date(r.checkInTime).toDateString()!==new Date(now).toDateString()){
              next[r.number]={...r,status:"overnight"}; changed=true; continue;
            }
            // overtime detection
            const overBy=(now-r.checkInTime)/3600000-baseHours(r);
            if(overBy>0&&r.status==="occupied"){
              next[r.number]={...r,status:"overtime"}; changed=true;
              if(settings.overtimeAlert&&!alertedRef.current.has(r.number)){
                alertedRef.current.add(r.number);
                addToast(`⏰ Phòng ${r.number} đã quá giờ ${Math.floor(overBy*60)} phút`,"warn");
              }
            }
            // auto-checkout
            if(settings.autoCheckout&&r.status==="overtime"){
              const overMin=(now-r.checkInTime)/60000-baseHours(r)*60;
              if(overMin>=settings.graceMins){
                const bill=calcBill(r,now);
                setHistory(h=>[{id:now+Math.random(),room:r.number,guestName:r.guestName,cccd:r.cccd,dob:r.dob,rateType:r.rateType,checkInTime:r.checkInTime,checkOutTime:now,duration:now-r.checkInTime,roomBase:calcRoomBase(r,now),extras:[...(r.extras||[])],extrasTotal:calcExtras(r.extras),bill},...h]);
                setTotalRev(t=>t+bill);
                next[r.number]={...EMPTY_ROOM(r.number),status:"cleaning"}; changed=true;
                alertedRef.current.delete(r.number);
                addToast(`🤖 Phòng ${r.number} đã được tự động trả phòng`,"warn");
              }
            }
          }
        }
        return changed?next:prev;
      });
    },5000);
    return()=>clearInterval(id);
  },[settings,addToast]);

  const openRoom=(r)=>{setSel(r.number);setDetailTab("info");setModal(r.status==="available"?"checkin":"detail");};
  const closeModal=()=>{setSel(null);setModal(null);setForm(EMPTY_FORM);setShowQR(false);setQrScanned(false);setDetailTab("info");};

  const handleQR=useCallback((raw)=>{
    setShowQR(false);
    const p=parseCCCDQR(raw);
    if(p){setForm(f=>({...f,...p}));setQrScanned(true);addToast("✅ Đọc CCCD thành công","success");}
    else addToast("❌ Không nhận dạng được mã QR","error");
  },[addToast]);

  const doCheckIn=()=>{
    setRooms(p=>({...p,[sel]:{...p[sel],status:"occupied",checkInTime:Date.now(),rateType:form.rate,guestName:form.name.trim(),cccd:form.cccd.trim(),dob:form.dob,extras:[]}}));
    alertedRef.current.delete(sel);
    addToast(`✅ Nhận phòng ${sel} thành công`,"success");
    closeModal();
  };

  const addDrink=(drink)=>{
    setPopId(drink.id);setTimeout(()=>setPopId(null),250);
    setRooms(p=>{const r={...p[sel]},extras=[...r.extras],idx=extras.findIndex(e=>e.id===drink.id);
      if(idx>=0)extras[idx]={...extras[idx],qty:extras[idx].qty+1};else extras.push({...drink,qty:1});
      return{...p,[sel]:{...r,extras}};
    });
  };
  const removeDrink=(id)=>setRooms(p=>{const r={...p[sel]};const extras=r.extras.map(e=>e.id===id?{...e,qty:e.qty-1}:e).filter(e=>e.qty>0);return{...p,[sel]:{...r,extras}};});

  const doCheckout=(finalBill)=>{
    const now=Date.now(),r=rooms[sel];
    const rec={id:now,room:sel,guestName:r.guestName,cccd:r.cccd,dob:r.dob,rateType:r.rateType,checkInTime:r.checkInTime,checkOutTime:now,duration:now-r.checkInTime,roomBase:calcRoomBase(r,now),extras:[...(r.extras||[])],extrasTotal:calcExtras(r.extras),bill:finalBill};
    setHistory(h=>[rec,...h]);
    setTotalRev(t=>t+finalBill);
    setRooms(p=>({...p,[sel]:{...EMPTY_ROOM(sel),status:"cleaning"}}));
    alertedRef.current.delete(sel);
    addToast(`💰 Phòng ${sel} đã trả phòng: ${fmtMoney(finalBill)}`,"success");
    closeModal();
    setTimeout(()=>printInvoice(rec),300);
  };
  const doCleaned=()=>{setRooms(p=>({...p,[sel]:{...p[sel],status:"available"}}));addToast(`🧹 Phòng ${sel} sẵn sàng nhận khách`,"success");closeModal();};

  /* Computed */
  const now=tick, room=sel?rooms[sel]:null;
  const bill=room?calcBill(room,now):0, elapsed=room?.checkInTime?now-room.checkInTime:0;
  const curDate=new Date(now), weekend=isWeekend(now);
  const stats=Object.values(rooms).reduce((a,r)=>({...a,[r.status]:(a[r.status]||0)+1}),{});
  const activeRev=Object.values(rooms).filter(r=>["occupied","overnight","overtime"].includes(r.status)).reduce((s,r)=>s+calcBill(r,now),0);
  const todayHist=history.filter(h=>sameDay(h.checkOutTime,now));
  const t=theme==="dark";

  if(!loaded)return <div style={{minHeight:"100vh",background:t?"#0b0c16":"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{color:"#c9a455",fontSize:14}}>Đang tải dữ liệu…</div></div>;

  return(
    <div style={{minHeight:"100vh",background:t?"#0b0c16":"#f1f5f9",fontFamily:"'DM Sans',sans-serif",color:t?"#e2e8f0":"#1e293b",paddingBottom:70}}>
      <style>{buildCSS(t)}</style>

      {/* ═ HEADER ═ */}
      <header style={{background:t?"linear-gradient(180deg,#0f1020 0%,#0b0c16 100%)":"linear-gradient(180deg,#fff 0%,#f8fafc 100%)",borderBottom:`1px solid ${t?"rgba(201,164,84,0.12)":"rgba(0,0,0,0.07)"}`,padding:"14px 20px",position:"sticky",top:0,zIndex:50}}>
        <div style={{maxWidth:1100,margin:"0 auto"}}>
          <div className="header-row" style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:12}}>
            <div>
              <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:700,color:"#c9a455",letterSpacing:2,lineHeight:1}}>◆ QUẢN LÝ KHÁCH SẠN</div>
              <div style={{fontSize:11,color:weekend?"#f87171":t?"#64748b":"#94a3b8",marginTop:3}}>{DAY_NAMES[curDate.getDay()]} — {fmtDate(now)}{weekend&&" ⚡ Cuối tuần"}</div>
            </div>
            <div className="rev-cards" style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              <div className="rev-card" style={{padding:"7px 14px",background:"rgba(248,113,113,0.08)",borderRadius:10,border:"1px solid rgba(248,113,113,0.2)",textAlign:"center",minWidth:130}}>
                <div style={{fontSize:9,color:"#b45454",fontWeight:700,letterSpacing:.8}}>ĐANG CÓ KHÁCH</div>
                <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:700,color:"#f87171"}}>{fmtMoney(activeRev)}</div>
              </div>
              <div className="rev-card" style={{padding:"7px 14px",background:"rgba(201,164,84,0.07)",borderRadius:10,border:"1px solid rgba(201,164,84,0.2)",textAlign:"center",minWidth:130}}>
                <div style={{fontSize:9,color:"#a8834a",fontWeight:700,letterSpacing:.8}}>TỔNG DOANH THU</div>
                <div key={totalRev} className="count-up" style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:700,color:"#c9a455"}}>{fmtMoney(totalRev)}</div>
              </div>
              <div className="rev-card" style={{padding:"7px 14px",background:"rgba(74,222,128,0.06)",borderRadius:10,border:"1px solid rgba(74,222,128,0.15)",textAlign:"center",minWidth:70}}>
                <div style={{fontSize:9,color:"#166534",fontWeight:700,letterSpacing:.8}}>HÔM NAY</div>
                <div key={todayHist.length} className="count-up" style={{fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:700,color:"#4ade80"}}>{todayHist.length}</div>
              </div>
              <div style={{fontFamily:"'DM Mono',monospace",fontSize:22,fontWeight:500,color:t?"#e2e8f0":"#1e293b",letterSpacing:2,alignSelf:"center"}}>
                {curDate.toLocaleTimeString("vi-VN",{hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false})}
              </div>
            </div>
          </div>
          <div className="stat-bar" style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:10}}>
            {Object.entries(STATUS_CFG).map(([key,cfg])=>(stats[key]>0||key==="available")&&(
              <div key={key} className="sp" style={{background:cfg.bg,border:`1px solid ${cfg.border}`}}>
                <span className={["occupied","overnight","overtime"].includes(key)?"pulse":""} style={{width:6,height:6,borderRadius:"50%",background:cfg.dot,display:"inline-block"}}/>
                <span style={{color:t?"#94a3b8":"#64748b"}}>{cfg.label}</span>
                <span style={{color:cfg.color,fontWeight:700,fontSize:13}}>{stats[key]||0}</span>
              </div>
            ))}
            <div className="sp" style={{marginLeft:"auto",color:t?"#64748b":"#94a3b8"}}>16 phòng</div>
          </div>
        </div>
      </header>

      {/* ═ CONTENT ═ */}
      <main style={{maxWidth:1100,margin:"0 auto"}}>
        {/* ── ROOMS VIEW ── */}
        {view==="rooms"&&(
          <div style={{padding:"20px 20px 0"}}>
            {Object.entries(FLOORS).map(([floor,nums])=>(
              <section key={floor} style={{marginBottom:28}}>
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
                  <div style={{width:3,height:20,background:"linear-gradient(#c9a455,#7c5c20)",borderRadius:2}}/>
                  <span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:18,fontWeight:700,color:"#c9a455",letterSpacing:1}}>Tầng {floor}</span>
                  <div style={{flex:1,height:1,background:`${t?"rgba(201,164,84,0.1)":"rgba(0,0,0,0.06)"}`}}/>
                </div>
                <div className="room-grid" style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  {nums.map(num=>{
                    const r=rooms[num],cfg=STATUS_CFG[r.status]||STATUS_CFG.available;
                    const active=["occupied","overnight","overtime"].includes(r.status);
                    const extCnt=(r.extras||[]).reduce((s,e)=>s+e.qty,0);
                    return(
                      <div key={num} className="room-card" onClick={()=>openRoom(r)}
                        style={{width:140,padding:"14px 13px",background:t?cfg.bg:"#fff",border:`1.5px solid ${cfg.border}`,borderRadius:12,boxShadow:t?`0 0 20px ${cfg.glow}`:`0 2px 8px rgba(0,0,0,0.06)`,position:"relative",overflow:"hidden"}}>
                        <div style={{position:"absolute",top:0,right:0,width:36,height:36,background:`radial-gradient(circle at top right,${cfg.color}18,transparent 70%)`,pointerEvents:"none"}}/>
                        {r.rateType&&<div style={{position:"absolute",top:8,right:8,fontSize:8,padding:"1px 6px",borderRadius:6,background:t?"rgba(255,255,255,0.08)":"rgba(0,0,0,0.06)",color:t?"#64748b":"#94a3b8",fontWeight:700}}>{r.rateType==="hourly"?"GIỜ":"NGÀY"}</div>}
                        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:26,fontWeight:700,color:t?"#e2e8f0":"#1e293b",lineHeight:1,marginBottom:8}}>{num}</div>
                        <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:6}}>
                          <span className={active?"pulse":""} style={{width:6,height:6,borderRadius:"50%",background:cfg.dot,display:"inline-block",flexShrink:0}}/>
                          <span style={{fontSize:11,color:cfg.color,fontWeight:600}}>{cfg.label}</span>
                        </div>
                        {active&&r.checkInTime&&(
                          <>
                            <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:t?"#64748b":"#94a3b8",letterSpacing:.5}}>{fmtDur(now-r.checkInTime)}</div>
                            <div style={{fontSize:12,color:"#c9a455",fontWeight:700,marginTop:4}}>{fmtMoney(calcBill(r,now))}</div>
                            {extCnt>0&&<div style={{fontSize:9,color:"#f59e0b",marginTop:2}}>🍺 {extCnt} món</div>}
                            {r.guestName&&<div style={{fontSize:9,color:t?"#475569":"#94a3b8",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.guestName}</div>}
                          </>
                        )}
                        {r.status==="cleaning" &&<div style={{fontSize:10,color:"#fbbf24",marginTop:2}}>Cần dọn dẹp</div>}
                        {r.status==="available"&&<div style={{fontSize:10,color:"#4ade80", marginTop:2}}>Sẵn sàng</div>}
                      </div>
                    );
                  })}
                </div>
              </section>
            ))}

            {/* History table */}
            <div style={{background:t?"rgba(255,255,255,0.02)":"#fff",border:`1px solid ${t?"rgba(255,255,255,0.08)":"#e2e8f0"}`,borderRadius:14,overflow:"hidden",marginBottom:16}}>
              <div onClick={()=>setHistOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"14px 20px",cursor:"pointer",borderBottom:histOpen&&todayHist.length>0?`1px solid ${t?"rgba(255,255,255,0.06)":"#f1f5f9"}`:"none",userSelect:"none"}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:3,height:18,background:"linear-gradient(#4ade80,#16a34a)",borderRadius:2}}/>
                  <span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:700,color:"#4ade80"}}>Lịch sử hôm nay</span>
                  {todayHist.length>0&&<span style={{fontSize:11,padding:"1px 8px",borderRadius:20,background:"rgba(74,222,128,0.12)",color:"#4ade80",fontWeight:600}}>{todayHist.length}</span>}
                </div>
                <div style={{display:"flex",gap:14,alignItems:"center"}}>
                  {todayHist.length>0&&<span style={{fontSize:13,color:"#c9a455",fontWeight:600}}>{fmtMoney(todayHist.reduce((s,h)=>s+h.bill,0))}</span>}
                  <span style={{color:t?"#475569":"#94a3b8",transform:histOpen?"rotate(180deg)":"none",display:"inline-block",transition:"transform .2s"}}>▾</span>
                </div>
              </div>
              {histOpen&&(todayHist.length===0?
                <div style={{padding:"32px",textAlign:"center",color:t?"#334155":"#94a3b8",fontSize:13}}>🏨 Chưa có lượt trả phòng hôm nay</div>:
                <div style={{overflowX:"auto"}}>
                  <table className="ht">
                    <thead><tr><th>PHÒNG</th><th>KHÁCH</th><th>CCCD</th><th>NGÀY SINH</th><th>VÀO</th><th>RA</th><th>T.GIAN</th><th>LOẠI</th><th>PHÒNG</th><th>ĐỒ UỐNG</th><th style={{textAlign:"right"}}>TỔNG</th><th></th></tr></thead>
                    <tbody>
                      {todayHist.map((h,i)=>(
                        <tr key={h.id} style={i===0?{animation:"rowAppear .35s ease"}:{}}>
                          <td><span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:700}}>{h.room}</span></td>
                          <td>{h.guestName||"—"}</td>
                          <td><span style={{fontFamily:"'DM Mono',monospace",fontSize:10,letterSpacing:.5}}>{h.cccd?h.cccd.replace(/(\d{4})(\d{4})(\d{4})/,"$1 $2 $3"):"—"}</span></td>
                          <td>{fmtDob(h.dob)}</td>
                          <td><span style={{fontFamily:"'DM Mono',monospace",fontSize:11}}>{fmtTime(h.checkInTime)}</span></td>
                          <td><span style={{fontFamily:"'DM Mono',monospace",fontSize:11}}>{fmtTime(h.checkOutTime)}</span></td>
                          <td>{fmtDurS(h.duration)}</td>
                          <td><span className="badge" style={{background:h.rateType==="hourly"?"rgba(96,165,250,0.1)":"rgba(167,139,250,0.1)",color:h.rateType==="hourly"?"#60a5fa":"#a78bfa",border:`1px solid ${h.rateType==="hourly"?"rgba(96,165,250,0.25)":"rgba(167,139,250,0.25)"}`}}>{h.rateType==="hourly"?"Ngắn giờ":"Ngày"}</span></td>
                          <td>{fmtMoney(h.roomBase)}</td>
                          <td>{h.extras?.length>0?<div>{h.extras.map(e=><div key={e.id} style={{fontSize:10,color:"#f59e0b",whiteSpace:"nowrap"}}>{e.icon}×{e.qty}</div>)}<div style={{fontSize:10,color:t?"#64748b":"#94a3b8"}}>{fmtMoney(h.extrasTotal)}</div></div>:"—"}</td>
                          <td style={{textAlign:"right"}}><span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:14,fontWeight:700,color:"#c9a455"}}>{fmtMoney(h.bill)}</span></td>
                          <td><button onClick={(e)=>{e.stopPropagation();printInvoice(h);}} style={{padding:"4px 8px",borderRadius:6,background:"rgba(201,164,84,0.1)",border:"1px solid rgba(201,164,84,0.2)",color:"#c9a455",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>🖨️</button></td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td colSpan={10} style={{padding:"10px",borderTop:`1px solid ${t?"rgba(201,164,84,0.2)":"rgba(201,164,84,0.3)"}`,fontSize:11,color:t?"#64748b":"#94a3b8",fontWeight:600}}>TỔNG {todayHist.length} LƯỢT</td>
                        <td style={{padding:"10px",borderTop:`1px solid ${t?"rgba(201,164,84,0.2)":"rgba(201,164,84,0.3)"}`,textAlign:"right"}}><span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:700,color:"#c9a455"}}>{fmtMoney(todayHist.reduce((s,h)=>s+h.bill,0))}</span></td>
                        <td style={{borderTop:`1px solid ${t?"rgba(201,164,84,0.2)":"rgba(201,164,84,0.3)"}`}}></td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              )}
            </div>

            {/* Price bar */}
            <div style={{padding:"14px 18px",background:t?"rgba(201,164,84,0.04)":"rgba(201,164,84,0.05)",border:`1px solid ${t?"rgba(201,164,84,0.12)":"rgba(201,164,84,0.2)"}`,borderRadius:12,marginBottom:20}}>
              <div style={{display:"flex",gap:18,flexWrap:"wrap",fontSize:12,color:t?"#64748b":"#94a3b8"}}>
                {[["⏱","Ngắn giờ","120k/2h","#e2e8f0"],["🌙","Ngày thường","220k/12h","#e2e8f0"],["🎉","Cuối tuần","250k/12h","#f87171"],["➕","Phụ thu","+20k/h vượt","#fbbf24"],["🍺","Đồ uống","10k/món","#f59e0b"]].map(([i,l,v,c])=>(
                  <span key={l} style={{display:"flex",alignItems:"center",gap:5}}>{i} {l}: <strong style={{color:c}}>{v}</strong></span>
                ))}
              </div>
            </div>
          </div>
        )}
        {view==="stats"&&<StatsView history={history} theme={theme}/>}
        {view==="settings"&&<SettingsView theme={theme} setTheme={setTheme} settings={settings} setSettings={setSettings} history={history} rooms={rooms} totalRevenue={totalRev} onLogout={onLogout} currentUser={currentUser} addToast={addToast}/>}
      </main>

      {/* ═ NAV BAR ═ */}
      <nav style={{position:"fixed",bottom:0,left:0,right:0,background:t?"rgba(10,11,22,0.95)":"rgba(255,255,255,0.95)",borderTop:`1px solid ${t?"rgba(255,255,255,0.08)":"rgba(0,0,0,0.07)"}`,backdropFilter:"blur(12px)",display:"flex",padding:"8px 20px",zIndex:50,justifyContent:"center",gap:4}}>
        {[{id:"rooms",ico:"🏨",label:"Phòng"},{id:"stats",ico:"📊",label:"Thống kê"},{id:"settings",ico:"⚙️",label:"Cài đặt"}].map(tab=>(
          <button key={tab.id} className={`nav-btn${view===tab.id?" act":""}`} onClick={()=>setView(tab.id)}>
            <span style={{fontSize:20}}>{tab.ico}</span>
            {tab.label}
          </button>
        ))}
      </nav>

      {/* ═ MODAL ═ */}
      {modal&&room&&(
        <div className="modal-bg" onClick={closeModal} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.72)",backdropFilter:"blur(6px)",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div className="modal-box mi" onClick={e=>e.stopPropagation()} style={{background:t?"#10111d":"#fff",border:`1px solid ${t?"rgba(201,164,84,0.18)":"rgba(0,0,0,0.1)"}`,borderRadius:18,padding:26,width:"100%",maxWidth:420,boxShadow:"0 32px 64px rgba(0,0,0,0.5)"}}>

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}>
              <div>
                <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:26,fontWeight:700,color:"#c9a455",lineHeight:1}}>Phòng {room.number}</div>
                <div style={{fontSize:12,color:(STATUS_CFG[room.status]||STATUS_CFG.available).color,marginTop:4,fontWeight:500}}>● {(STATUS_CFG[room.status]||STATUS_CFG.available).label}</div>
              </div>
              <button onClick={closeModal} style={{background:t?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.05)",border:`1px solid ${t?"rgba(255,255,255,0.1)":"rgba(0,0,0,0.08)"}`,color:t?"#64748b":"#94a3b8",cursor:"pointer",width:30,height:30,borderRadius:"50%",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}}>✕</button>
            </div>

            {/* CHECK-IN */}
            {modal==="checkin"&&(
              <>
                {showQR?<div style={{marginBottom:14}}><QRScanner onResult={handleQR} onClose={()=>setShowQR(false)}/></div>:(
                  <>
                    <button className="qb" onClick={()=>setShowQR(true)} style={{marginBottom:12}}>
                      <span style={{fontSize:20}}>📷</span> Quét mã QR Căn Cước Công Dân
                    </button>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
                      <div style={{flex:1,height:1,background:t?"rgba(255,255,255,0.07)":"#f1f5f9"}}/>
                      <span style={{fontSize:11,color:t?"#334155":"#cbd5e1"}}>hoặc nhập thủ công</span>
                      <div style={{flex:1,height:1,background:t?"rgba(255,255,255,0.07)":"#f1f5f9"}}/>
                    </div>
                    {qrScanned&&<div style={{marginBottom:10,padding:"8px 12px",background:"rgba(74,222,128,0.08)",border:"1px solid rgba(74,222,128,0.2)",borderRadius:8,fontSize:12,color:"#4ade80"}}>✅ Đã quét CCCD — kiểm tra thông tin bên dưới</div>}
                    <div style={{fontSize:11,color:"#c9a455",fontWeight:600,letterSpacing:1,marginBottom:8}}>THÔNG TIN KHÁCH</div>
                    <div style={{padding:12,background:t?"rgba(255,255,255,0.02)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.06)":"#e2e8f0"}`,borderRadius:10,marginBottom:12,display:"flex",flexDirection:"column",gap:10}}>
                      <div><label className="fl">TÊN KHÁCH</label><input className="fi" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Nguyễn Văn A..."/></div>
                      <div>
                        <label className="fl">SỐ CĂN CƯỚC CÔNG DÂN</label>
                        <input className="fi" value={form.cccd} onChange={e=>{const v=e.target.value.replace(/\D/g,"");if(v.length<=12)setForm(f=>({...f,cccd:v}));}} placeholder="12 chữ số..." style={{fontFamily:"'DM Mono',monospace",letterSpacing:2}}/>
                        {form.cccd.length>0&&form.cccd.length<12&&<div style={{fontSize:11,color:"#f59e0b",marginTop:4}}>⚠ {form.cccd.length}/12 chữ số</div>}
                        {form.cccd.length===12&&<div style={{fontSize:11,color:"#4ade80",marginTop:4}}>✓ Hợp lệ</div>}
                      </div>
                      <div>
                        <label className="fl">NGÀY THÁNG NĂM SINH</label>
                        <input className="fi" type="date" value={form.dob} onChange={e=>setForm(f=>({...f,dob:e.target.value}))} max={new Date().toISOString().split("T")[0]}/>
                        {form.dob&&<div style={{fontSize:11,color:t?"#64748b":"#94a3b8",marginTop:4}}>{fmtDob(form.dob)} · {new Date().getFullYear()-parseInt(form.dob.split("-")[0])} tuổi</div>}
                      </div>
                    </div>
                    <div style={{marginBottom:12}}>
                      <div style={{fontSize:11,color:"#c9a455",fontWeight:600,letterSpacing:1,marginBottom:8}}>LOẠI PHÒNG</div>
                      <div style={{display:"flex",gap:8}}>
                        {[{v:"hourly",l:"Ngắn giờ",s:"120k/2h",i:"⏱"},{v:"daily",l:"Theo ngày",s:weekend?"250k/12h":"220k/12h",i:"🌙"}].map(opt=>{
                          const act=form.rate===opt.v;
                          return(<div key={opt.v} onClick={()=>setForm(f=>({...f,rate:opt.v}))} style={{flex:1,padding:"10px 8px",borderRadius:10,cursor:"pointer",textAlign:"center",border:`1px solid ${act?"#c9a455":t?"rgba(255,255,255,0.08)":"#e2e8f0"}`,background:act?"rgba(201,164,84,0.1)":t?"rgba(255,255,255,0.03)":"#f8fafc",transition:"all .15s"}}>
                            <div style={{fontSize:18,marginBottom:2}}>{opt.i}</div>
                            <div style={{fontSize:12,fontWeight:600,color:act?"#c9a455":t?"#94a3b8":"#64748b"}}>{opt.l}</div>
                            <div style={{fontSize:11,color:act?"#a8834a":t?"#475569":"#94a3b8",marginTop:1}}>{opt.s}</div>
                          </div>);
                        })}
                      </div>
                    </div>
                    <div style={{fontSize:12,color:t?"#475569":"#94a3b8",marginBottom:14,padding:"9px 12px",background:t?"rgba(255,255,255,0.03)":"#f8fafc",borderRadius:8,lineHeight:1.7}}>
                      <div>🕐 Vào: <strong style={{color:t?"#94a3b8":"#475569"}}>{curDate.toLocaleTimeString("vi-VN",{hour:"2-digit",minute:"2-digit",hour12:false})}</strong> — {DAY_NAMES[curDate.getDay()]}</div>
                      {weekend&&<div style={{color:"#f87171"}}>⚡ Giá cuối tuần</div>}
                    </div>
                    <button className="btn" onClick={doCheckIn} style={{background:"linear-gradient(135deg,#c9a455,#8c6a28)",color:"#0a0b12"}}>✓ Nhận phòng</button>
                  </>
                )}
              </>
            )}

            {/* DETAIL */}
            {modal==="detail"&&(
              <>
                {["occupied","overnight","overtime"].includes(room.status)&&(
                  <>
                    <div style={{display:"flex",gap:6,marginBottom:14}}>
                      {[{id:"info",l:"📋 Thông tin"},{id:"drinks",l:"🍺 Đồ uống"}].map(tb=>(
                        <button key={tb.id} className={`tb${detailTab===tb.id?" act":""}`} onClick={()=>setDetailTab(tb.id)}>{tb.l}</button>
                      ))}
                      {(room.extras||[]).reduce((s,e)=>s+e.qty,0)>0&&(
                        <span style={{marginLeft:"auto",fontSize:11,color:"#f59e0b",alignSelf:"center",fontWeight:600}}>{(room.extras||[]).reduce((s,e)=>s+e.qty,0)} món</span>
                      )}
                    </div>

                    {detailTab==="info"&&(
                      <>
                        <div style={{padding:"12px 14px",background:t?"rgba(255,255,255,0.03)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:10,marginBottom:12}}>
                          <div style={{fontSize:10,color:"#c9a455",fontWeight:700,letterSpacing:.8,marginBottom:10}}>THÔNG TIN KHÁCH</div>
                          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                            <div style={{gridColumn:"1/-1"}}><div style={{fontSize:10,color:t?"#475569":"#94a3b8",fontWeight:600}}>TÊN</div><div style={{fontSize:13,marginTop:2}}>{room.guestName||"—"}</div></div>
                            <div><div style={{fontSize:10,color:t?"#475569":"#94a3b8",fontWeight:600}}>CCCD</div><div style={{fontSize:12,marginTop:2,fontFamily:"'DM Mono',monospace",letterSpacing:.5}}>{room.cccd?room.cccd.replace(/(\d{4})(\d{4})(\d{4})/,"$1 $2 $3"):"—"}</div></div>
                            <div><div style={{fontSize:10,color:t?"#475569":"#94a3b8",fontWeight:600}}>NGÀY SINH</div><div style={{fontSize:12,marginTop:2}}>{fmtDob(room.dob)}</div></div>
                          </div>
                        </div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
                          {[{l:"THỜI GIAN",v:fmtDur(elapsed),m:true},{l:"LOẠI",v:room.rateType==="hourly"?"Ngắn giờ":"Theo ngày"},{l:"GIỜ VÀO",v:fmtTime(room.checkInTime),m:true},{l:"NGÀY VÀO",v:fmtDate(room.checkInTime)}].map(({l,v,m})=>(
                            <div key={l} style={{padding:"10px 12px",background:t?"rgba(255,255,255,0.03)":"#f8fafc",borderRadius:8}}>
                              <div style={{fontSize:9,color:t?"#475569":"#94a3b8",fontWeight:700,letterSpacing:.5,marginBottom:3}}>{l}</div>
                              <div style={{fontSize:13,fontFamily:m?"'DM Mono',monospace":"inherit",letterSpacing:m?1:0}}>{v}</div>
                            </div>
                          ))}
                        </div>
                      </>
                    )}

                    {detailTab==="drinks"&&(
                      <div style={{marginBottom:12}}>
                        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
                          {DRINKS.map(drink=>{
                            const qty=(room.extras||[]).find(e=>e.id===drink.id)?.qty||0;
                            return(<div key={drink.id} className="db" onClick={()=>addDrink(drink)} style={{position:"relative",animation:popId===drink.id?"pop .25s ease":undefined}}>
                              <span style={{fontSize:20}}>{drink.icon}</span>
                              <span style={{fontSize:10,fontWeight:600,color:t?"#94a3b8":"#64748b",textAlign:"center"}}>{drink.name}</span>
                              <span style={{fontSize:9,color:t?"#475569":"#94a3b8"}}>10.000đ</span>
                              {qty>0&&<div style={{position:"absolute",top:-5,right:-5,width:16,height:16,borderRadius:"50%",background:"#c9a455",color:"#0a0b12",fontSize:9,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>{qty}</div>}
                            </div>);
                          })}
                        </div>
                        {(room.extras||[]).length>0?(
                          <div style={{background:t?"rgba(255,255,255,0.02)":"#f8fafc",border:`1px solid ${t?"rgba(255,255,255,0.07)":"#e2e8f0"}`,borderRadius:10,overflow:"hidden"}}>
                            {(room.extras||[]).map(e=>(
                              <div key={e.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",borderBottom:`1px solid ${t?"rgba(255,255,255,0.04)":"#f1f5f9"}`}}>
                                <span style={{fontSize:13}}>{e.icon} {e.name}</span>
                                <div style={{display:"flex",alignItems:"center",gap:8}}>
                                  <span style={{fontSize:11,color:t?"#64748b":"#94a3b8"}}>{fmtMoney(e.qty*e.price)}</span>
                                  <div style={{display:"flex",alignItems:"center",gap:5}}>
                                    <button className="tb" onClick={()=>removeDrink(e.id)} style={{width:22,height:22,padding:0,fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:"50%"}}>−</button>
                                    <span style={{fontSize:13,fontWeight:600,minWidth:16,textAlign:"center"}}>{e.qty}</span>
                                    <button className="tb" onClick={()=>addDrink(e)} style={{width:22,height:22,padding:0,fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:"50%"}}>+</button>
                                  </div>
                                </div>
                              </div>
                            ))}
                            <div style={{display:"flex",justifyContent:"space-between",padding:"8px 12px",background:"rgba(245,158,11,0.05)"}}>
                              <span style={{fontSize:12,color:"#a8834a",fontWeight:600}}>Tổng đồ uống</span>
                              <span style={{fontSize:13,color:"#f59e0b",fontWeight:700}}>{fmtMoney(calcExtras(room.extras||[]))}</span>
                            </div>
                          </div>
                        ):<div style={{textAlign:"center",padding:"16px",color:t?"#334155":"#94a3b8",fontSize:13}}>Nhấn để thêm đồ uống</div>}
                      </div>
                    )}

                    {/* Bill */}
                    <div style={{padding:14,background:"rgba(201,164,84,0.07)",border:"1px solid rgba(201,164,84,0.2)",borderRadius:12,marginBottom:14}}>
                      <div style={{fontSize:10,color:"#a8834a",marginBottom:8,fontWeight:700,letterSpacing:1,textAlign:"center"}}>TỔNG TIỀN</div>
                      <div style={{display:"flex",flexDirection:"column",gap:4,marginBottom:8,fontSize:12,color:t?"#64748b":"#94a3b8"}}>
                        <div style={{display:"flex",justifyContent:"space-between"}}><span>Tiền phòng</span><span style={{color:t?"#94a3b8":"#475569"}}>{fmtMoney(calcRoomBase(room,now))}</span></div>
                        {(room.extras||[]).length>0&&<div style={{display:"flex",justifyContent:"space-between"}}><span>🍺 Đồ uống</span><span style={{color:"#f59e0b"}}>{fmtMoney(calcExtras(room.extras))}</span></div>}
                        <div style={{height:1,background:"rgba(201,164,84,0.2)",margin:"3px 0"}}/>
                      </div>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:30,fontWeight:700,color:"#c9a455"}}>{fmtMoney(bill)}</div>
                        <div style={{fontSize:11,color:t?"#64748b":"#94a3b8",marginTop:2}}>{fmtDur(elapsed)}</div>
                      </div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:8}}>
                      <button className="btn" onClick={()=>doCheckout(bill)} style={{background:"linear-gradient(135deg,#ef4444,#991b1b)",color:"#fff"}}>Trả phòng — {fmtMoney(bill)}</button>
                      <button onClick={()=>{const r=rooms[sel];if(r.checkInTime)printInvoice({...r,checkOutTime:Date.now(),duration:Date.now()-r.checkInTime,roomBase:calcRoomBase(r,now),extrasTotal:calcExtras(r.extras),bill});}} style={{padding:"13px 14px",borderRadius:10,background:"rgba(201,164,84,0.1)",border:"1px solid rgba(201,164,84,0.2)",color:"#c9a455",cursor:"pointer",fontSize:16,fontFamily:"inherit"}}>🖨️</button>
                    </div>
                  </>
                )}
                {room.status==="cleaning"&&(
                  <><div style={{textAlign:"center",padding:"24px 0 20px"}}>
                    <div style={{fontSize:44,marginBottom:12}}>🧹</div>
                    <div style={{fontSize:15,color:"#fbbf24",fontWeight:600,marginBottom:6}}>Đang dọn dẹp</div>
                    <div style={{fontSize:13,color:t?"#475569":"#94a3b8"}}>Xác nhận sau khi dọn xong để mở phòng</div>
                  </div>
                  <button className="btn" onClick={doCleaned} style={{background:"linear-gradient(135deg,#22c55e,#15803d)",color:"#fff"}}>✓ Dọn xong — Mở phòng</button></>
                )}
              </>
            )}
          </div>
        </div>
      )}

      <ToastContainer toasts={toasts}/>
    </div>
  );
}

/* ═══════════════════ ROOT ═══════════════════ */
export default function App(){
  const [user,  setUser]  = useState(null);
  const [theme, setTheme] = useState("dark");
  const handleTheme=(v)=>{setTheme(v);DB.save("theme",v);};
  useEffect(()=>{DB.load("theme","dark").then(setTheme);DB.load("session",null).then(u=>{if(u)setUser(u);});}, []);
  const login=(u)=>{setUser(u);DB.save("session",u);};
  const logout=()=>{setUser(null);DB.save("session",null);};
  if(!user)return <LoginScreen onLogin={login} theme={theme}/>;
  return <HotelApp currentUser={user} onLogout={logout} theme={theme} setTheme={handleTheme}/>;
}
