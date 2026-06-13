import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

// ─── DESIGN TOKENS ────────────────────────────────────────────────────────────
const C = {
  bg:"#0F1E35",sidebar:"#111827",primary:"#00B4D8",primaryL:"#E0F7FF",
  success:"#10B97B",successL:"#D1FAE5",warning:"#FFD700",warningL:"#FFF9C4",
  danger:"#EF4444",dangerL:"#FEE2E2",purple:"#FF00FF",purpleL:"#FFE6FF",
  text:"#0F172A",muted:"#64748B",border:"#E2E8F0",surface:"#F8FAFC",
};

// ─── SEED DATA ────────────────────────────────────────────────────────────────
const seedProducts = [
  {id:1,sku:"PRD-001",name:"Laptop Asus VivoBook",category:"Elektronik",stock:45,minStock:10,price:8500000,unit:"Unit",warehouse:"Gudang A"},
  {id:2,sku:"PRD-002",name:"Mouse Logitech M235",category:"Aksesoris",stock:120,minStock:20,price:185000,unit:"Unit",warehouse:"Gudang A"},
  {id:3,sku:"PRD-003",name:"Keyboard Mechanical",category:"Aksesoris",stock:8,minStock:15,price:650000,unit:"Unit",warehouse:"Gudang B"},
  {id:4,sku:"PRD-004",name:'Monitor LG 24"',category:"Elektronik",stock:23,minStock:5,price:3200000,unit:"Unit",warehouse:"Gudang A"},
  {id:5,sku:"PRD-005",name:"Headset Sony",category:"Audio",stock:56,minStock:10,price:420000,unit:"Unit",warehouse:"Gudang B"},
];
const seedCustomers = [
  {id:1,name:"PT Maju Bersama",contact:"Budi Santoso",phone:"021-55512345",email:"budi@majubersama.co.id",address:"Jl. Sudirman No. 12, Jakarta"},
  {id:2,name:"CV Teknologi Nusantara",contact:"Siti Rahayu",phone:"022-87654321",email:"siti@teknusantara.com",address:"Jl. Asia Afrika No. 5, Bandung"},
  {id:3,name:"Toko Elektronik Jaya",contact:"Ahmad Fauzi",phone:"031-99887766",email:"ahmad@elektronikjaya.id",address:"Jl. Darmo No. 88, Surabaya"},
];
const seedOrders = [
  {id:"SO-2024-001",date:"2024-06-01",customer:"PT Maju Bersama",status:"delivered",items:[{productId:1,name:"Laptop Asus VivoBook",qty:3,price:8500000},{productId:2,name:"Mouse Logitech M235",qty:3,price:185000}],total:26055000,invoiceId:"INV-2024-001"},
  {id:"SO-2024-002",date:"2024-06-05",customer:"CV Teknologi Nusantara",status:"processing",items:[{productId:4,name:'Monitor LG 24"',qty:5,price:3200000}],total:16000000,invoiceId:null},
  {id:"SO-2024-003",date:"2024-06-08",customer:"Toko Elektronik Jaya",status:"pending",items:[{productId:5,name:"Headset Sony",qty:10,price:420000},{productId:2,name:"Mouse Logitech M235",qty:15,price:185000}],total:6975000,invoiceId:null},
];
const seedInvoices = [
  {id:"INV-2024-001",orderId:"SO-2024-001",date:"2024-06-01",dueDate:"2024-06-15",customer:"PT Maju Bersama",total:26055000,tax:2605500,grandTotal:28660500,status:"paid",paidDate:"2024-06-10"},
];
const seedMovements = [
  {id:1,date:"2024-06-01",type:"out",subtype:"penjualan",product:"Laptop Asus VivoBook",productId:1,qty:3,ref:"SO-2024-001",note:"Pengiriman order",before:48,after:45},
  {id:2,date:"2024-06-01",type:"out",subtype:"penjualan",product:"Mouse Logitech M235",productId:2,qty:3,ref:"SO-2024-001",note:"Pengiriman order",before:123,after:120},
  {id:3,date:"2024-06-03",type:"in",subtype:"pembelian",product:"Keyboard Mechanical",productId:3,qty:20,ref:"PO-2024-001",note:"Restock dari supplier",before:8,after:28},
  {id:4,date:"2024-06-05",type:"out",subtype:"penjualan",product:'Monitor LG 24"',productId:4,qty:5,ref:"SO-2024-002",note:"Pengiriman order",before:28,after:23},
];
const seedCashFlows = [
  {id:"CF-M-001",date:"2024-06-02",type:"out",category:"Operasional",description:"Gaji karyawan Juni",amount:15000000,ref:"PAYROLL-06"},
  {id:"CF-M-002",date:"2024-06-04",type:"out",category:"Logistik",description:"Biaya pengiriman",amount:1200000,ref:"SHIP-001"},
  {id:"CF-M-003",date:"2024-06-07",type:"out",category:"Utilitas",description:"Listrik & internet",amount:2500000,ref:"UTIL-06"},
  {id:"CF-M-004",date:"2024-06-09",type:"in",category:"Lain-lain",description:"Pengembalian deposit gudang",amount:3000000,ref:"DEP-RET-01"},
];

const seedEmployees = [
  {id:1,nik:"EMP-001",name:"Budi Hartono",dept:"Teknologi",position:"Staff IT",status:"aktif",joinDate:"2022-01-15",gajiPokok:6500000,tunjMakan:500000,tunjTransport:400000,tunjJabatan:0,potonganBPJS:195000,potonganPPh:0,bank:"BCA",noRek:"1234567890"},
  {id:2,nik:"EMP-002",name:"Sari Dewi",dept:"Keuangan",position:"Akuntan",status:"aktif",joinDate:"2021-06-01",gajiPokok:7000000,tunjMakan:500000,tunjTransport:400000,tunjJabatan:500000,potonganBPJS:210000,potonganPPh:100000,bank:"Mandiri",noRek:"0987654321"},
  {id:3,nik:"EMP-003",name:"Rizky Pratama",dept:"Gudang",position:"Staff Gudang",status:"aktif",joinDate:"2023-03-10",gajiPokok:4500000,tunjMakan:500000,tunjTransport:300000,tunjJabatan:0,potonganBPJS:135000,potonganPPh:0,bank:"BNI",noRek:"1122334455"},
  {id:4,nik:"EMP-004",name:"Maya Indah",dept:"Penjualan",position:"Sales Manager",status:"aktif",joinDate:"2020-08-20",gajiPokok:9000000,tunjMakan:600000,tunjTransport:500000,tunjJabatan:1500000,potonganBPJS:270000,potonganPPh:350000,bank:"BCA",noRek:"5566778899"},
  {id:5,nik:"EMP-005",name:"Doni Saputra",dept:"Gudang",position:"Kepala Gudang",status:"aktif",joinDate:"2019-11-05",gajiPokok:8000000,tunjMakan:600000,tunjTransport:500000,tunjJabatan:1000000,potonganBPJS:240000,potonganPPh:200000,bank:"Mandiri",noRek:"6677889900"},
];
const seedPayrolls = [
  {id:"PAY-2024-06-001",employeeId:1,employeeName:"Budi Hartono",month:"2024-06",gajiPokok:6500000,tunjMakan:500000,tunjTransport:400000,tunjJabatan:0,lemburJam:0,lemburAmount:0,potonganBPJS:195000,potonganPPh:0,potonganLain:0,totalTunjangan:900000,totalPotongan:195000,takehome:7205000,status:"dibayar",payDate:"2024-06-28"},
  {id:"PAY-2024-06-002",employeeId:2,employeeName:"Sari Dewi",month:"2024-06",gajiPokok:7000000,tunjMakan:500000,tunjTransport:400000,tunjJabatan:500000,lemburJam:0,lemburAmount:0,potonganBPJS:210000,potonganPPh:100000,potonganLain:0,totalTunjangan:1400000,totalPotongan:310000,takehome:8090000,status:"dibayar",payDate:"2024-06-28"},
];

const seedSuppliers = [
  {id:1,name:"PT Sumber Makmur",contact:"Agus Wijaya",phone:"021-44433322",email:"agus@sumbermakmur.co.id",address:"Jl. Gatot Subroto No. 45, Jakarta",terms:"30 hari"},
  {id:2,name:"CV Indo Elektronik",contact:"Hendra Kusuma",phone:"022-77665544",email:"hendra@indoelektronik.com",address:"Jl. Braga No. 12, Bandung",terms:"14 hari"},
  {id:3,name:"UD Jaya Teknik",contact:"Wawan Setiawan",phone:"031-88776655",email:"wawan@jayateknik.id",address:"Jl. Pemuda No. 33, Surabaya",terms:"COD"},
];
const seedWarehouses = [
  {id:1,name:"Gudang A",location:"Lantai 1, Gedung Utama",pic:"Doni Saputra",capacity:500,used:193},
  {id:2,name:"Gudang B",location:"Lantai 2, Gedung Utama",pic:"Rizky Pratama",capacity:300,used:64},
  {id:3,name:"Gudang C",location:"Gedung Belakang",pic:"",capacity:200,used:0},
];
const seedPOs = [
  {id:"PO-2024-001",date:"2024-06-03",supplierId:1,supplierName:"PT Sumber Makmur",status:"diterima",items:[{productId:3,name:"Keyboard Mechanical",qty:20,price:450000}],subtotal:9000000,tax:900000,total:9900000,notes:"Restock rutin",receivedDate:"2024-06-05"},
  {id:"PO-2024-002",date:"2024-06-10",supplierId:2,supplierName:"CV Indo Elektronik",status:"dipesan",items:[{productId:1,name:"Laptop Asus VivoBook",qty:10,price:7200000}],subtotal:72000000,tax:7200000,total:79200000,notes:"",receivedDate:null},
];
const seedReturns = [
  {id:"RTN-2024-001",date:"2024-06-07",type:"penjualan",refId:"SO-2024-001",customer:"PT Maju Bersama",items:[{productId:2,name:"Mouse Logitech M235",qty:1,price:185000}],total:185000,reason:"Barang cacat",status:"disetujui",notes:"Diterima kembali ke gudang"},
];
const seedDiscounts = [
  {id:1,name:"Diskon Member Gold",type:"persen",value:10,minOrder:5000000,customerId:null,productId:null,active:true,desc:"Diskon 10% untuk order di atas 5 juta"},
  {id:2,name:"Harga Khusus PT Maju",type:"persen",value:5,minOrder:0,customerId:1,productId:null,active:true,desc:"Diskon 5% khusus PT Maju Bersama"},
  {id:3,name:"Promo Laptop Juni",type:"nominal",value:500000,minOrder:0,customerId:null,productId:1,active:true,desc:"Potongan Rp500rb untuk Laptop VivoBook"},
];

const MONTHS=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
const seedSalesChart=[
  {month:"2024-01",revenue:42000000,orders:8,expenses:28000000},
  {month:"2024-02",revenue:55000000,orders:11,expenses:31000000},
  {month:"2024-03",revenue:48000000,orders:9,expenses:29000000},
  {month:"2024-04",revenue:63000000,orders:14,expenses:35000000},
  {month:"2024-05",revenue:71000000,orders:16,expenses:38000000},
  {month:"2024-06",revenue:49030000,orders:10,expenses:36200000},
];

const seedAttendance=[
  {id:1,employeeId:1,date:"2024-06-10",checkIn:"08:02",checkOut:"17:05",status:"hadir",note:""},
  {id:2,employeeId:2,date:"2024-06-10",checkIn:"07:55",checkOut:"17:00",status:"hadir",note:""},
  {id:3,employeeId:3,date:"2024-06-10",checkIn:"08:30",checkOut:"17:00",status:"terlambat",note:"Macet"},
  {id:4,employeeId:4,date:"2024-06-10",checkIn:"",checkOut:"",status:"izin",note:"Sakit"},
  {id:5,employeeId:5,date:"2024-06-10",checkIn:"08:00",checkOut:"17:00",status:"hadir",note:""},
  {id:6,employeeId:1,date:"2024-06-11",checkIn:"08:05",checkOut:"17:10",status:"hadir",note:""},
  {id:7,employeeId:2,date:"2024-06-11",checkIn:"08:00",checkOut:"17:00",status:"hadir",note:""},
];

const seedUsers=[
  {id:1,username:"admin",name:"Administrator",role:"admin",email:"admin@gks.co.id",active:true,lastLogin:"2024-06-10 08:30"},
  {id:2,username:"kasir1",name:"Sari Dewi",role:"kasir",email:"sari@gks.co.id",active:true,lastLogin:"2024-06-10 09:15"},
  {id:3,username:"gudang1",name:"Doni Saputra",role:"gudang",email:"doni@gks.co.id",active:true,lastLogin:"2024-06-09 14:00"},
  {id:4,username:"manager1",name:"Maya Indah",role:"manager",email:"maya@gks.co.id",active:true,lastLogin:"2024-06-10 10:00"},
];

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt  = n => new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",minimumFractionDigits:0}).format(n);
const fmtN = n => new Intl.NumberFormat("id-ID").format(n);
const fmtR = n => Number(Number(n).toFixed(0));
const today = () => new Date().toISOString().slice(0,10);
const uid   = () => Date.now() + Math.random();

const card  = {background:"#FFF",borderRadius:12,padding:20,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",border:"1px solid #E2E8F0"};
const tdS   = {padding:"10px 12px",verticalAlign:"middle"};
const btn   = bg=>({background:bg,color:"#FFF",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontWeight:600,fontSize:13});
const btnSm = bg=>({background:bg,color:"#FFF",border:"none",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontWeight:600,fontSize:12});
const inp   = {width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #E2E8F0",fontSize:13,boxSizing:"border-box",outline:"none",fontFamily:"inherit"};
const lbl   = {fontSize:12,fontWeight:600,color:C.muted,display:"block",marginBottom:6};
const sel   = {...inp};

// ─── REUSABLE UI ──────────────────────────────────────────────────────────────
const Badge = ({status})=>{
  const m={
    pending:{l:"Pending",bg:C.warningL,c:"#92400E"},processing:{l:"Diproses",bg:"#DBEAFE",c:"#1E40AF"},
    delivered:{l:"Terkirim",bg:C.successL,c:"#065F46"},cancelled:{l:"Dibatalkan",bg:C.dangerL,c:"#991B1B"},
    paid:{l:"Lunas",bg:C.successL,c:"#065F46"},unpaid:{l:"Belum Bayar",bg:C.dangerL,c:"#991B1B"},
    overdue:{l:"Jatuh Tempo",bg:C.warningL,c:"#92400E"},
    in:{l:"Masuk",bg:C.successL,c:"#065F46"},out:{l:"Keluar",bg:C.dangerL,c:"#991B1B"},
    opname:{l:"Opname",bg:C.purpleL,c:"#5B21B6"},
    aman:{l:"Aman",bg:C.successL,c:"#065F46"},kritis:{l:"Kritis",bg:C.dangerL,c:"#991B1B"},
    aktif:{l:"Aktif",bg:C.successL,c:"#065F46"},nonaktif:{l:"Non-Aktif",bg:"#F1F5F9",c:"#475569"},
    dibayar:{l:"Dibayar",bg:C.successL,c:"#065F46"},proses:{l:"Diproses",bg:"#DBEAFE",c:"#1E40AF"},belum:{l:"Belum Bayar",bg:C.dangerL,c:"#991B1B"},
    dipesan:{l:"Dipesan",bg:"#DBEAFE",c:"#1E40AF"},diterima:{l:"Diterima",bg:C.successL,c:"#065F46"},dibatalkan:{l:"Dibatalkan",bg:C.dangerL,c:"#991B1B"},sebagian:{l:"Sebagian",bg:C.warningL,c:"#92400E"},
    disetujui:{l:"Disetujui",bg:C.successL,c:"#065F46"},ditolak:{l:"Ditolak",bg:C.dangerL,c:"#991B1B"},menunggu:{l:"Menunggu",bg:C.warningL,c:"#92400E"},
    selesai:{l:"Selesai",bg:C.successL,c:"#065F46"},transfer:{l:"Transfer",bg:"#DBEAFE",c:"#1E40AF"},
    hadir:{l:"Hadir",bg:C.successL,c:"#065F46"},terlambat:{l:"Terlambat",bg:C.warningL,c:"#92400E"},
    izin:{l:"Izin",bg:"#DBEAFE",c:"#1E40AF"},alpha:{l:"Alpha",bg:C.dangerL,c:"#991B1B"},sakit:{l:"Sakit",bg:"#EDE9FE",c:"#5B21B6"},
    admin:{l:"Admin",bg:"#1E3A5F",c:"#FFF"},kasir:{l:"Kasir",bg:C.primaryL,c:C.primary},
    gudang:{l:"Gudang",bg:C.successL,c:"#065F46"},manager:{l:"Manager",bg:C.purpleL,c:C.purple},
  };
  const s=m[status]||{l:status,bg:"#F1F5F9",c:"#475569"};
  return <span style={{background:s.bg,color:s.c,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:600}}>{s.l}</span>;
};

function Modal({children,onClose,wide=false}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:24}}>
      <div style={{background:"#FFF",borderRadius:16,padding:28,maxWidth:wide?820:620,width:"100%",maxHeight:"88vh",overflow:"auto",position:"relative",boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
        <button onClick={onClose} style={{position:"absolute",top:16,right:16,background:"#F1F5F9",border:"none",borderRadius:"50%",width:32,height:32,cursor:"pointer",fontSize:14,color:"#64748B"}}>✕</button>
        {children}
      </div>
    </div>
  );
}

function ConfirmModal({msg,onConfirm,onCancel}){
  return(
    <Modal onClose={onCancel}>
      <div style={{textAlign:"center",padding:"16px 0"}}>
        <div style={{fontSize:44,marginBottom:12}}>🗑️</div>
        <div style={{fontWeight:700,fontSize:17,marginBottom:8}}>Hapus Data?</div>
        <div style={{fontSize:13,color:C.muted,marginBottom:24}}>{msg}</div>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel} style={{...btn("#94A3B8")}}>Batal</button>
          <button onClick={onConfirm} style={{...btn(C.danger)}}>Ya, Hapus</button>
        </div>
      </div>
    </Modal>
  );
}

// Field helper
function F({label,children}){return <div><label style={lbl}>{label}</label>{children}</div>;}

// Action buttons – edit + delete
function ActionBtns({onEdit,onDelete,extraBtns}){
  return(
    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
      {extraBtns}
      <button onClick={onEdit} style={{...btnSm("#64748B"),display:"flex",alignItems:"center",gap:4}}>✏️ Edit</button>
      <button onClick={onDelete} style={{...btnSm(C.danger),display:"flex",alignItems:"center",gap:4}}>🗑️</button>
    </div>
  );
}

// ─── EXPORT UTILITIES ─────────────────────────────────────────────────────────
function exportExcel(sheets){
  const wb=XLSX.utils.book_new();
  // Cover sheet with branding
  const coverData=[
    ["GKS — Global Karunia Supply"],
    ["Digital Printing Distributor"],
    [""],
    ["GKS ERP System — Laporan Export"],
    [`Tanggal Export: ${new Date().toLocaleString("id-ID")}`],
    [""],
    ["Dokumen ini digenerate otomatis oleh GKS ERP System"],
  ];
  const wsCover=XLSX.utils.aoa_to_sheet(coverData);
  wsCover["!cols"]=[{wch:50}];
  XLSX.utils.book_append_sheet(wb,wsCover,"Info");
  sheets.forEach(({name,cols,rows})=>{
    // Header rows with branding
    const headerRows=[
      ["GKS — Global Karunia Supply | Digital Printing Distributor"],
      [`Laporan: ${name} | Diekspor: ${new Date().toLocaleString("id-ID")}`],
      [],
      cols,
      ...rows,
    ];
    const ws=XLSX.utils.aoa_to_sheet(headerRows);
    ws["!cols"]=cols.map(()=>({wch:22}));
    XLSX.utils.book_append_sheet(wb,ws,name);
  });
  XLSX.writeFile(wb,`GKS_ERP_${today()}.xlsx`);
}
function exportPDF(title,html){
  const LOGO="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHgAAAB4CAIAAAC2BqGFAAABCGlDQ1BJQ0MgUHJvZmlsZQAAeJxjYGA8wQAELAYMDLl5JUVB7k4KEZFRCuwPGBiBEAwSk4sLGHADoKpv1yBqL+viUYcLcKakFicD6Q9ArFIEtBxopAiQLZIOYWuA2EkQtg2IXV5SUAJkB4DYRSFBzkB2CpCtkY7ETkJiJxcUgdT3ANk2uTmlyQh3M/Ck5oUGA2kOIJZhKGYIYnBncAL5H6IkfxEDg8VXBgbmCQixpJkMDNtbGRgkbiHEVBYwMPC3MDBsO48QQ4RJQWJRIliIBYiZ0tIYGD4tZ2DgjWRgEL7AwMAVDQsIHG5TALvNnSEfCNMZchhSgSKeDHkMyQx6QJYRgwGDIYMZAKbWPz9HbOBQAABBgUlEQVR42u29d7glVZU2vtbeu/KJN3fOdCI0oaFFgk2QDKIDigqjYMAwpvnAMc84hhFHP3UUBQlKEnFIkhUkiQTJdDedc/eN555Qp/Lee/3+OLdzoBtaxd9nPfXc59y6de6p89aqFd4VNrquC//Y/vIb+wcE/wD6/1eb+PsTDWQAoIkA6B9A77sNcctPQCAdyZCALO4wxgkICEYQJ/oH0HuJLDJAAAIgDVqBUqAVaNKamGkc0XGsgcYLlT8moc8ZAgNgHJggzhD5lje+yXDHN4XXsRlcpUimmMqRw6ZBjsu8ovbyzMsr0zjbOOn/8A9YYP1YXn+dvAuziMIAmw0dNjAIKU1G/qEhwDCACYA3C+h/U6ARgTEgoizFJEUAsm3s7IFxU2ncFBo3hXrGqrYuypfIzZFt22RedceoE18CAHh2Grz/3IGqGWGUQNxkjYYYHqCB9Wz9Sly7AteuoMH1EEYaAE0DDAsYgv5bIv43Uh2MASBkCSQpMGBdY2DageqAuTTjID1msmrvBicHnIGSlMQUBRA0sVqJkuiPzXWHqtkC2BPR4vqaCnmutl3IFXXnaDnrYOScNGEYsuoA610tFr/MFjyDS16kvvWkNJgmmBYAgFb/D0g040AaogA0YFcPHXI0HTFfzTxE9owHL4dKQaOqN66ldStx/SrRt9aoDJhh3YiallJAaVHkTu443ULr3spdQ0kFmRlzntle4hZkW5fsHkdjJ+K4qThmApTKwE2Imrx/o1jyInv6YXzucdq4hhDA8YAzUPqv6bf8BYBG3PkTyjhoBWGIhoDZc/VxZ8m5R6sxU8D2IPRp/Qpa8LxY/Ly7bmmhXimC9gxuWZZpW5KYBpAAiiCVWQYSAAQISxgMkSMxrQ1EmcVxnESprBFr5NvCsVPl9INg/7k4YRrkCpDGYuNq/sKT7A93wIt/oiQhx0Zh7BTuXX2DNxPQiKQVItvmYltSHAToenTEcXTyu7ODjtDlDkxTWrWYnnrEeulP5f7VXUwVPddwcwlgM80azTAIozhO0jRRUmqtiAg0ASK23DoGiIwxwQQ3DMOxbdexC3kvbxoWgoqCRrM5oHm1fWx8wBE473iYNpNsFxt1Y9Gz/P7fwBMPUKMGjouc0yZlgggEoJVmjO1buPcp0JoANLdcncaaCDkDQESkoAmGxeafnp35gWzmweB5ODyg//y4eOTu9tULR3NdKJW07Vab0VCtUW80kjgimSERgkYAhsAYMkRExtiWUJZIa02aSGutiQiQADUyLoRh2flCvr1Y6Ch4PI2btWpvSoNjp8ujT8N586lrNMSRsXyBuPtG+N2tMmpyN8cINGitCQBd2w7jBEAzxvaVetmXQDMg89gz8ZhTYdFzyd036DSFLIE0Y/Pmq3M/lsw5khyXD2xQD93pPPrbcfX+7nIRvMJAM+zrHwp8X8mUackROGOGYZiWaRiGaVqcc2QMAQGAiFqxCxEhMkRs4ayVTrMkS7M0TdI0U1prAsU4F6br5rq6O3qKOYyaQ5XGOrcYHHkqnvxOGj0RskS8+lzp+qsaT9+VGORhnnP68DnFk94GDzwKP7+lkcqMIaN9gfU+ApoxiiNr/7nJpd+LusYaSVz82X8077pJjJ1E5386OfZMVSzyoX79+1u9h26bIptt7W1N4qs39leHqzpLGGmO2hSG63m245imSQRKySRJ0jTJskxJqZQiIK1p61gckTGOQgjDMAzDtCxLCIGIMsuiKAzDME0zRaCRkzBLpcLEMaOLJtQHKyvIbBx9Jjv1HBg9/qrrCoXHf//V/i8uGHjxXSf1XP+dyCnXw0rhgi8UbntwyHP41h/6t3fvkIgEJ2ECY6QlppKfcm72wUvSMZN41GT3/Mq87ZppWX3ClMkb6s6fl62KmgGjjIGyhenlip7nMc7TJAmawVA8mGWZ1hqRMcYZE8gNZtjIGCDbQoRpBaQ1ySRTcRxo3SDSDNEwDNtxHdftyhc16SgIms1mkgX+UPxSpWK53pSJE47ubOt7/LZX/njH9EM+fGrvZ7yeU/bvPPJ7yy+riCsUa4LmhqlNA4kA32SqA4E0csM5/f3xYW+1BvvT9p7gkLeC44qFz6vrvj9p7cIJ48cPRtnKNevDwDdIMdCO4xZLJSGMOI5834+jUGuNnHNhc8Nm3ERmIIpW3NhSHTtuBAREBK0wXSkVa5lImZCUiOjYTq6Qtx1HK9Wo14OgScBSZLbjTBo3bnQ+17tubVs4/fMTv35M+YhUwZPJo8NH/9u4uX/+8xPjL/vFUDNQnO8bk7jvdDQy0Bqi0DzurPhDlyajJ4kkVrdfXbzz2gPHjk5NZ9GyFYHfEKCRVC5XKJSKpKnRqDebDa00F5YwHWY4yC1kApCN8ESbOaOtsd18d7e+04g48jcNWmmd6iySaaCylCG6uVyxWOSC+3Xf9+samARmu96MaZOFKVcv3vBe98OfmfKFgrY26oH/O/z5ny/9hQWOKbgi/WaSaGSkFYsjOOej8QWf1oV2Y91y/YMvTu5b2jNhyore/r7ePgMJZOblcuVyOctkvVYNwxA5M8wcN3OMW8AEAtAI/0mwCxHeUzcTEQA0SZCpSv0sDZSSju2UymXTsuq1qt/wiQtJrLu9Y/LEMWuHVvcMTLxsxs8O9WY3EvrJ4He+u/orDLjBTE37IJLkhmHsg2BPZUgEn/zP+H2fQscTj91jfPNf5npMtHU9v2Bx6De4Tk3Bu3tGWbY1NDhUrVaIwHQLptvBrSIKc5PfvVNpfX0bARACY9xkZk6YnhBGlsWNei2O41KxVCyVsySiLAnTbMNAZVLPRD46u3rR5UXsObR84FH5o6bmDnxo+J5IhyazCNQbvKQ3DDTjkMbMMOmLP4lPOoeTout+1H7Ntw+dtd/qqr94yXITMpBpW3tHqa2tVq0ODQ5qANstG26nMArABBDBPno8d4c448xwuOkJbsosqteqUsr2zk7LtuNmgyGtHxxCJWbMmnzTmmv7q5Wj2k880JlxWOltv6/cW5UVm9kaNL4BrN+Y6mAMkhjzbeqrl8cHH2VEgfrhFya98OiEmbOfW/hqGDSFyoQhurq74yipDA0SkGkVuV1EbgIRgN4Xkru3aQQEnamknsZ10rqtvSPn5QYGB5Ik1sy2LOvQg2Yv37j4gOpbfnzAdR3Meylc/IFFZ62OluVEXpL8WwDNOKUxy7XJ//hZctBb7NqQ+s7nZm5YXJo4+ennX2Ja6zQqFIvFYnFoaCgImqblCLuNG54GwL+sCO+JBgclIxXX0jiwbLurqysIgmq1wgw7Az7vwAP6go2d68ZddeAtY0Xngmj5+QtPXxMtd4WnXi/WrxdoZKAyxk31zWuiQ442axX66oenV9bkJkx65vmXbAY6Szo6u4Qw+vv6NCnTKXG7xFDQ3xjircHmRJlMG1lYA6LO7m6G2N/fx7kVKTps/9mVpN9eW755zn3jja6Xo6XnvDJ/KB10uK1el218XUAjAmlUmr780+iYU636kPrqR/arroZSz6tLltiCkUy7e0YlSVKpDAlhGE47N3O0E0ftb482MNRpmEVDWRqXy+25vNe3sQ8YixTMmDAhM0JzXemWOQ+MMzr/6D/z/gWnhDow0NCw1+LyeowhMgZhCJ/8z/ikc0Qc6m99akbfMih3L16+3OaESvWMHu37fq1aNWzPcruY6bx5BHkHS0mMm8xwGamg2VBKd3V1hUFDIA7U/a58R1quP7jid6d0nzPDnjTBm/3bwV9xNF6HYdl7oLmAps/O/Uj0vn/hpNT3Pz975YveuEkvvbLA5oSKukeNqlarzaZvOnnT64QRu4fw5t00Y5wbHqKOw0aSZp2d3WHY5ER91frknsmNQv8flz1+es+5BzgzTKPw+6Hf2tyhvRTqvQMamaBmg809Ovncd8By6Lr/mfT4HV1T93vq2RdsgaBUz6hRw8PDYRiYTtFwOwHw76H6YqSWgZsuQ53EfpKkXV1dYRgIxLWDg9PHTVuBi9b2rj+h8+RDvHlrstUv1J9xuLNXCmRvgGaM0pS3d2Vf/ansGms8ek/b1d+YdtCcJ5970UBNMh01Zky1Wg3DwHJLptsJIyExwt/HRgDAjBwyyOJmkqRd3d1Nv2FytrZ/aM602Q8P3VfIRs8rzjkkf+zDtft70w0mM/ecQd0LoAmRp6n+1+8khxxtrFlmfPOTh+4//dlFS3WaUJb0jBrj+41m0zedoul20ggT8/eC8silEpAQOQCdxr6UqrOzy6/VOBcbB4cOPWjOrctuekvupNnuxCnewbcP3AhAex7C7DHQnEOzyc54X/yej/Mkom9+aq7HVteC4aEhzJKOru4sS2vV4a00xt/l1rpuZrgIOgrryHip3OY3hhVgHKqJ+/Xc++pvz+h+70x7csLwkaEHbLGnynrPihwRIYn5mAnZe/+FLEvdds3k3qWp661bu07oNF8sCiEqQ0OG7RlOByAAaPg73giAhNNmOvl6bVhp1dbWwWTWO9Sf+lgb1XvZiv8EBhf3fObg0uGhDBjyfQY0ImKW6Qs+l46bIha+ULjtqlGTp764cInJtBBGsVjs7+sThmHaHcA4EPz9SvQWqJEZTodhOYMD/a7r2o5tIy1YtmxcecIt4VUPDf1xlOF+ceJlDNkeeq57oDoYpyBgh8+PP3AJJ6m/d+ncvLFiaDhs1FGmo0aPrlQqWZaaXhc3nDd5peHeOrKMGToNwijq7u5u1GvIMAjl1BnjHl360Bnd79/PmbI6Wf1C45k98fb2QKJJMcNU7/6YLpb0H347ce2CzPH6evuYytraO+I4CYOm6ZS46dFOUUYcKa1jDFq5KMQ9U5eb39PaOe76arc9c2R/TUvVKknjDDijkTds/Q6tmXCFW07T2Pf9zq4uyNKhekUFor9j7S/XXe0x+MSYL5SNcqbT1/ys15JoxiEI2PFnRf/0Ed6omP/3C/uP7n5p2UpIY0OIcrmtv69XmLbhdm5/zxCBcQQArSDLIMtAZihTUBKIABE4HykM25mHhIAZZSGFGWUpJSmlCSVEJHAnSU6GLKIopjijNN20xzrmjLeKqXc4HzgHTZBmGKeYpJBkPJOgCAmQM+CcOENAICImHNRps1kvldukkiDlUKM5a7+pj6166O3ld89wxw2q+p+GH3lNq/hayVkt0fWysz4Ajqtvv2Zq1hhK7ND3hZadPWOq1QoBCbsNt2aLWvKbxpjEZJhUKFChBG4ehAGkKYkxbLJGHYMmSQmGAZbdCoW3RlmS7BZdk63JEhQSEpBAPqSGVyTLOIitvVeGLFDN2c6B7bwkSQEgIGjSOZ5bGC+syCGx7fmcQxhhlGAxp8aPyjpL2rWBCKNUN5qsWudVH4OYacUMk1xbM0BhlzMZVwaHOro6N6xfJ5NosL8pxqlr1/70W9O/dkH3x27pu7YuqwLFbtxq8Rri3PRh/hnZzEP4wEbn97+ZuN+0Pzy3UID2cjkAajQatlvmhke0KQHBOGQxJIkaP0kfdmQ2+2A1apzOF8E0ARkAkdIsiVijznrX82WLjAXP46qlSACGuRlrhixQwdG5c340/pt1CQxBEZQE3FZ78MOrP5gXhc06SiCvZJX3tL3vO2O/JcBsHZUE7QbcVvvDv677LNv2OWMMqnU8aL/kgjOD4w4PJo2Kc64CzgAAFKUJ1ptiw4BYutZ6frH1zELr5WW20sQM27QLYXM4S4vFYqler69cv27+0Yfd0/frC4KLZ7jj3tH13ivXfd8yrN0Q1rsFmjQahj753eB56vZrJ8vm+nohavoWqFK5NDgwyIXBrSJtdkAZh9DXHd3ZO89P3zpfFcqkFMgUlAKloHURiMp2lJeDcRNh3rEsDMxVS61rfsjWrALT3FquNelEUqxTjlyRSrSZ6Wxrf4Yjr8jK+9rOv2zsZZnWEWWIKEm2C/uGyl2fW/dphtxAY7OUMQaNgC7959qXP15xixkkABmAApKq9RyaBnR2qM6eZM4hwbkcdc08+LxxyzcYjkXCKhhJWKlURo8e7TcaSRqvXzvAxsrr113x7ZlffXfXRTf3XpVRisB3RTmw3YlzHMGsQ7ODjmDDg84jv23r7Fy6fKWJKpcvSKmiKDTtAnJzJBHFGQQNddDc6Bs/Dk99pxQGNeoQ+pClICUoNQKi1pBlEEXQbECjppVM5hyux02GLNnRSCKOWNLN+9YoD8vhC9o++N1xl6VaSdAMmSLdYdi/rt752XWfFswwcUuIzDnUfPy3D9S+9eU+l2WyijoGJUEq0AQEoDVICVmMWRPTGsoqSSmRayAAIkBT2MUsTaIoKpVKAvSKtevacx33xbeujGpzvFnz206OVMx37VPv1utQWp/wDl3u0H9+bFy93weeBAGSLhSL9WqVC4OZhRH4GMNmMzv8qOjz307LbVCvAmngDACBccgXIF9A22WGhbYD+QIUS+DmgHNQCsImZOkeuiItDc6RD2fDF7V/6LIx34qVVEAMUJJqF+YNw7d9bv1nTGYKEJtJH8agGeJhM6OvfayiqqAVCE4AyG0QZeA5YBZjDooCGG1kFMk0CQEUbS6eQSDNjZww7Vqt6uVyDClJk6Ch4o7avf135gX9U9cHOPDd2EOxS8cnS7CzJzv0aExTfOSuzrbikvX9jDLHcUlTEIZ2roTMANKtuFFNmJJ84osZIiQpcAEATGtyPdb0zacf4EsW4FAfTzNlGLpUhjET5bSZctJ+ulAEtncteAzYUDb0oc6PfmP010IlCYABKlLtwvzl8C1fWH+pxW0OfGtqjSHECb7/tJB7SlaBCyANzIGVq3LX3+MtXmX4geBCt5fU9PHZIbOSQ2bG7aMyJ9HbMZfCKkT+QJpl+Vxe+c01GzZMmdZ117JfXTD2n99SOGZabuay4FWH7ZzVE7vOuqZ06FFq7FRctbh91UIc3V1bvp6DKpZKjUadccaNwlbqiOLzL5b5IjYbxDkAgFbk5u3nnrCu+xmuX4UEwJAQDQLSCglM29ZjJ6Tz5mennI2GsWeRDnLkVTn8sc5P/PvoLwUyI0QEVKTaDPPaoV99YeO/OdzZDmUAUBpcW8+dFZEEZKAJ0IKVq+23fbR73XpbWBoRgUBp1KRNAWO7s+MOj9799gC3iRY1MzwhzHqt1t7e7vsNv+kLNX61veTZ+kvHth309razFvsLkLOdJvXFrvIOyJg6/DiwHXrq4dFCD/ihlpktTCGMZtM3zNyIdmYMwyCbc3h20FwKfOAcAFBrdPPWUw/b3/93BEa5wuY8VstyEgJozdasspe9aj32AACAZb821giRHP5cz5e/OOoSX0pAREBNqk2YVw1e/+WNX3a5i4DboYwAmsB1oLOkUBEiaI3CoQefdNZtsMeNyeIEqUUb4MjJg8PG1bcbv7o3bztkmTRS40jEmOBWLgyrHR0dlmVnSTpYadidxv2D/3t8x0EnlE6/ct1/7yp7y3aqNyhNWWc3m3UgixrGS0/mi8W+gUFGysvl4jjSSnEztyXW0FoeOZ+E2HyTyDD4UJ91zY+ACXKcTf1retOuRmyjZUOxjP292N8L4jWEGgESFX2067NfGX2JL7MWLIpUGzd/NvTLL238kitcBNy5J0uAoAFa8REwJEjhoBmp46gNAyJKmKQtKBOBYVBHiSwLtNrOcqAwPABsNoNcPs+I+iuVglN8On5sMFEHeAdP9WYmOmY7Q3WnQDPM0nDa3LBzjlzb29a/iiw3bAYcted5vu9zYaIwR0h9rahQVNNmUpaNGDTSYDvGY79nA/1g2aDU1rEiMAaMj+yt+nrLBsvaPYHOkTeVPjZ37JdGfa0h1WZZLhvmj4eu/erGL3vCa8U1OyWIOMMg4gNVJN56VoEiOGJOcP/lve9+uz+2OyNN1ToOVpkfoFLIGGjaCaFApBm3hGE3mw3bcQTDKA5RmYPOhgX+gh7Tmls4SmmNOwtHdxrR6pCszxy9uGPyt2981FcA1SBUMnWEwRiLo9CwC4ibzGCW6q7Rsq0D5CagkWGaigXPkyFg63BRSUhiAEDahjAf+TqWA7st3NRAeVFqlRwiIBGVhfmjgZ9/Y+N/5I0CbOqh3VVqKEzwsReceUf51ABoNX4kcMy8+jFH1psVY9UG89VV1kvLzGdftV9eZgwMCcvCnKPl9hJNxIQw3TgYIk2WZSZROjzs8xI8WXvkhI4D5xWP/cXGn+z0SsTOSCRtWsaFh6+Z3fZNfx381nlrrd7LSDpuMUlSrTUznBGtgQhKUalEtgNpOlI/xzkLfBzsBy5GgEMEKancrmcdgFIR7qirGFv0AjQawMRucoyKVOsGadI5bnyv7/JvDny93ejQpHefUlIacg5cfUfhI+9qlEqxDEEIQCQVAADk3OyA2dkBc4JzESFmazeY9/7Ju+bW3HOL3VJeb3sDEYiYcABYFIWu6zXDuFqvj+8svzjwdKjwIO/wstkeSH/HcFzsSLhECUweixMmoByynlrseZ65ZqPPAWzXCZshco7c2vxcIZG2HGJ8S3oQEZMY02iL34aIWZKNmxz+6zd0EuG2LjMBMM5zX/qYqAyRY8Cuy+u3PAYICmiiPaHMyimlBhi7B5oIbItWbTQ+8OXOG77Xnyul5IPSgAgMgRRQ2KoQJs7U+NHRxedHH3hH47+uaP/2Lws5e9tSdCLgBhdGFIbltjaO4AeBxUavgWW9STDWGjvJmfp842mDGdupHraDfgYlYdYUzHXJgXW6t+IgZnGcMsZMw4zikAsbt5c72ok53Z42RFCS/Dr4DfIb1PQpaLZ2aP3Ue1KHRy2PggGLlDyndNrlE67gJFJI+WvxvUpB0aN7/pQ7/gNj73u4CIKLMnAPWmygJgRE3lIpKcgqMyH993/ru/SCet1HzretTUAuhBMnMedCCCZlKhNq2NVlweJ2wWY5B5GmHRldtrOkGc6cSpCjRSsglYU0i7TMTMMggCzLuLABtvQqETIWRajkFpiIyLLB8raXTc4pl4dcDnJ5MK0RMnjzvifJYQAHjZaaZsgGZfL2/DFXTLjSIJFQyl8rpSQ1lPL0ygrz7E/1HPuhsf/1k64nnikMVS20uCgBz1OrlB4RBNegQFXhixdVp09KowQZbo0QMmEpqZVWpmmRkkGQcA+WNBeYDKZ5M/coMmxhM3MKAafFK4EJz2/6SGSapswkac2EuYVDJgLOsV7FJAJhUotoVlJ7eerqwY1raDNPxBgLfPPlZynLUCvdPVq2d4JUe5jz0qTzwvhd/REJ2VmlE4eyRKAQKIZlckL+qCsmXHXxmo8klFho7aYwDgGUAtcGAPbMAu/x513Xou52td/4dM709JhDwuOPCCxXUUwtGZASnHJ60rzwBzeZrrVNXzNyg5CyNDUtC4O40QyLnWJlYwkATLFn7DS/tb0oaQ1C4OSxAApXrOeWZYVRjKAN08yyFBCRmbi58ogIhGCVATE0RJsdYSIyTHXgIaTkSIeE1mRaYtVS72ufzn37Uu/LH3eeepzZLuxxtSABcQBf1z+25sMP+U+0iRFCsoX18fm3XjnxKousWMccdk3rMOCs5cpTztGdJbItGKqJh5/1Lvtl6R2fHT3v/AkvLXLBBq0362OaPinbQeMTMMEYS5LUNE0GFESxZVjr5KpIwRhrYk7kFcjtci5s+3uuKe/h2M4MIlrXL2xbxHGKAIZhpGmGTADybbQ84+g3+ZKXwTBHLhAZRmF61Am6ZwwkEWzScIQMTAtMG0xbC75bf2ynQg0Wc1JKP7r64gf9x9u3xXq+95afT7rKRiemnVNoiOCH6IetBAoQgVRABIaggqfbS7pcohdftb/58zbkbOu2g5yTIW7ndRIyxpiRpYkQJkNI0tg0nAr1+5I6jO6iUZY7yNB2+SeQEksF7ChrCKFaNxlikqWMoRAiyxLODIbbJ59IcPOJh1mWAts0LUZmsq09ueizBIhhuCVIQQbAWv487WWmHBEUacFMhdlH1nzkPv+RbeRapcd4866eeLWDbqyj7bBGhDSDT7+nMX9uVG1gtc4yiZyBGMmmEQJyJGFCM+Kgt7mwMOJEDHd8PLiQMuOcMYYyy0CBzxv1LCjyQpvoULR7iUbQCkoFyuVU4lMjMAC1lIohMMaUlMA5bRf2aA2Oyxa9aD3zOHpFpjIAIMYoDNPDjky+8B01cSpGATbr0Gxg08dmHfwayvR19O8hgCYtmKFRfmzNR+9t/KFdWC3ZEcirMj3KO/zqide4mIt1tLUOQYQ4gTOO9e+7dt2dP+x972n1jmJWD2Goyip1HG7wSh17K9xz1CfPGwbcRhhXbzR36AwjAIbMkEoBImdMa6UUpCyqyYrLWdno0C1ztStjiABaUynHwcZGBePUtB2tlBIMAVApxUy2Ez6BCLhh3niFnHGAzJdG1AVjFDajAw/Jvv4jvuB5sexVHB4kJcFxdecoNfetEMd7S5BuNowGGhlkH1978U/GXX566YTKiG3kVZm+1Tvs6onXfGj1hU1qOszZbBsRodFEUOq0+bXTTmxUNprPLzFfWWqt6TP8EC2hJ4+RZ84Ppk+LKBzJbTEGELEnXnRME/S2URYCIGNaayBijFFGWmnJEz/1LYQiL+5Y2rKTyNB1NAiKItTEiRQQtTqxNdFO6RIgAtNifRvtH349uvRbynEgjKiVwY9Cybk8/Khk3rEjI3daqdskghY3QgRaA+d7TvwDgCYy0JSQfXzdxzRcfmbxxIrcjHVypHfoNZN+cdHqC33dsJmrN9FpBKA0Kp8E0+3l+MS3xSceD7C5k4YBxKDDkbufZWC0wx8eKjy9wM65Wm9PCG52BwCRESmlCDiEuskZuNzbse6Q7eBFY7mgwaK6r1IlSCutWzQMERHblQxqBV5OvPKC+81LzN4NUCoB56AVAKFS0PShUYNmAwIf/AbUq5AkQLrl4kChyBAx3THJQgREI/QOtX5uLdcCDYbsk2s/fkf9/nbDUqSINEc+nMWHOwdfM/GaIhRC7beEgwiKOeIecQ4aMEswq2M2jLIKsgGyDtkwqhgBUCpUCow26F/vfPp77YzvVMUhY4xIayLOGZCSioBDTVUFQlG07agX2Q7RCmlq3aidtEHTyPm089jLy/Oli5yvftK99XoRBixXBC9Plg2mCUIAE8AFGCY4Dng5yBeZ64ha1fndnbn/+ByuX0umBVtJDkduc7TRtpiw0bIZmmhup0MEcsb4v6z75G3Dd5eEaaJho8hxO9JwtHvwDVNvHiPGZpQhoGnATffln3/W0yhEiYwyGXkyHBImCAFCgGESN4k5JErA8+yxJwsnXzxq2VrT3eaithTWbte3i5vaGgFgbzIse22laESuHQ/jxPrl5eKBO9TBh8tZh6gxY6nQpi2LGEOlWJpC4LOhQb52hViykC1bxCoDyAVZdsvmtCyex7zHmo+dveJ8uYlI4sAqathhDtEWfamJBHICumTj52+q3rTZk0EASTrPXYubIEkT2BZcc2fhpvtyMyen8/aPD5sdT5+YjenIih5ZpkJGSrJmxDcM8eeXOHc9mvvdE7YCVvBoc1D1BmvdxE4dKRgZvrIHepP0NpegNSCQ57HKILvnf437bgfXIy9Htg1ckFIsiTGMMAxAJsQ4mBa4HhGBUoCgNz1EiNgrN67OVm0Tw4OwmSVBbUfpMWASskeDx7ZjoLQGE20TDakIAD2HNLEXFjtPv+wAgmvpfI4KnnQsYoxlEvyADfvYDBgyLHjEkZJsq7w7e22oaYTTfC0+ugVuw+eQYj6HnGeIHBmS1gDY0vpbq3lCANPZuR2zHCyWiAi1xiTBOBohoRFJcCqXNz1tm+4TAihlSy42hQfIWqPwtkBHQHo7L3ervxZZYetfNWmyI86V3jq3hFDwgCEBoNKgFNabrNpAAkAEzqCY020FDUBKb68444RvIW8ISGtERESlNSDjjIHCvF1QGppZbcdnYEeJpmYEkDHPIYGEjDFAIhqZHkCbPwsBNALPzvuQ6ujCzemVXQUbO9zP7R8Ly4aViz/zaPtsb/9oD2oGdy9WDjeW1Hv1KZe89fAwDXDEhO/IgwPB1imILdYWtxZRhqQ1++z3S2v7hGW0wACglvABaQJEzhko9HhOAQSq+VoSDYAMGgFAQjlXm0aCwBjnSmZAxDkjLTfzSa1qB/GrqwTui2poRFDqigxZReyT/iJNRL9xr7zd2VfNSmHCTLGFZCZSjDFA1FoxFMiRKyMv8gmBrxo7OhLbqw7OsdaAMGBuXhdd5QMTgsuUtNaci1Sq7eQR42g30I1E26SJgDHcPDMHERlia3QPArRYSCIIkYgIWyOviAiIIdLIbImty5dAE7WmHRERjqQQR+jTLU9NyjDlmggBEFETsZF/CwxRE7UmfGz+xE0TfHFneceRKGYzr6SV4pyTJqWJcc4EWdorme2R1MNyCHcYMca2+3+CU7VBg1UBOVYuJKC5aRqaSCllGAYpuX2nPNtpaTJDITSyRhA0moEiIMQozVp/YkJITVW/GaUZMwyFGCZpM4yl1hyFAIEa0zhlwAQKmWqdkUDBgXPgAoXOtEq1yc00zlSqBBqUUZZkrXNkIuMwyRIlUCBQlkWcEYCUMuaMsjRCVIK3jmslE9JScERUaRpxpgUHwWiHHcT28YbSWgphEBFpLYRADp7KFUW+ofxqVhG4fREe2+7OcY6NgHqHONgwukvFSWJZJgFmaWqaptaKtER8zQJvTOKEIR43/5gTTpjveZ5pmVMmT2pZjzCMioX86ae9fdqUyc1Gw7HtGdOnzZs3t1wqJkmkgWzH3m/aVM64knLMmFFdnR1KjngaSsqurs4xY0fXq7WpUyePHz9OpnFbe9ukSRO0UoAwbvzYQw85aNKkCVEcGqYxZcpkxjCfz08YPwEApk6batk2AEyZPJlzPm7c2LZyWcnMsZ1p06YyxrXeMpdlu30bs0xKK2mapsxSRWCblsySduzKC6xkA7VsmKOg3fPRjIFMadV6AEFTxmVpkriOQ4BZlpqmRaRBy9dEOctkV3fnrbfe9NnPfPKDH3z/GWecPGnyhF/dfL3rOkHQPODAWXfcecsF57/3xhuvPefcd40fP/bWW28++x1n3HLLTfOOOLzW33fhhee/8PTv3jb/6MZg79e//pWPfvTCer3OOeec1+v1T33qE5/4xMVnnHHKjTde7TpOHEWXffdbjz7ywJgxo5Ik+flVl3/5y/927bVXXHTRB3M577bbf207znHHHXv55T9kjN1++y1f/eq/CSFu+tUvi4X8d7/7zXPOfWd9cMNFF13w4tO/O+bYo5rNgHP+muaESBFlpmmmaUoAtm0nMh0jJjocNqTrmrLBgRPQa6ayaPFKBIXTJ5PKwkIuR8jSJBWGYAyVSnfPcDLG/IZ/4YUXGIbxrne97zv/9f0/PPS453oIWhhGlmSf+MRHV65a/a4zTvnxT6749Kc/Xsjn0jRZvGRpsZgngHypfPQxR91wy11nn30WASICE0II0Yr+kXG/6Z968ts+/omPXnDBh5999vk5hx1aLhcffuTxd77rnWEzdF33wQf/sGzZsu7ubq01Q+ScM84RwTBEZXjoqKOOfMtbjgiaAeMGY1xKaXptRx9z1A233H322WciQ3qtgikEIJUCkWlaaZoQYKGQS+N0sjWVA6yIl0hSO3YasJ2lsmjhCoIQZk8GwXzLsJkQaZYhgBCGlslrd7cRdbS3r12zLqgO//u/f/lrX/tiHCUyk4MDQzoZLuQL69auB8hWr1nTmlentTaEMTg06HreQYcePHnShD8+8fS8t8ztGDshDKOgGciwr9lsIiIipElaKBQMwxgaqkAannzSCQj08suvnHraia7nNur1004/5fDDD7vn7nsFNxiySmW46Te5YIgYNIPrf3njxRd/qDWAUxP5fvPguYdMnDThj0889ZZ5h0+aOCGO493rRgJQMuZccMGTJAUmPNdUAU33DkgIlgev7tSXFDsycVzAouUsGeTjJujuNj8B07IMGSQyyxzH9ZuBqSXgLp0wIrJd+ze/ufW6666+9EufHxgY7O7pAtRjxo/70hcvWbdu3f33PXDp5//P8qVfeN/573/44cfqDb+zs7O3t7dUKk2eOHHGjGlRFB04axpDfN97z03T9OSTToyj7yxevOTBhx7WSo8dO+b+3z32yssvP/TQfe97z/mnnnpStVof3dMzbuz4448/tlwqXXLpl2bPnvGTn/7fE084vdH0v/Klz++//6ylS1ckcTJlyuT7H3hwziFzTpg/L0vSUrFk29b573tPFIYHzpqGiOedd+7Xv/4tx3GUUrsUaFIqSyzL1kpJJbnhmRb34uK00bNqGS0KX0TEHet3t28WIgLBoeazd56I3TPks0/Big3dGpLAbxpCWLbT9BvCcpGbuwHasqyVK1c/8/Szhx8xt1avX3nlNf39A0MDQ8hQGMYNN9y8ZMnSI9/6lscf+9MPfvA/mnBgcGD8+HH33vu72269va2t7Yc/uPwXP//5KwuXakXPPf+C3/Q9z9uwfsOKlatM04yjaPmyZVdf9Ysoirp6ul986ZVvf/u/b7np5g19fX4jePnlBatWrfrTn54CwFdfXfKHhx4+7LBDVixf+f3v/49S2XCl+tJLr/z5z8+vWbPh+RderAxVlixdFsfJj39yxS+uvOqVRUuzLF2+YtWu1TQBMtBpElYLxaLW2g/CfL5cLBujqlM+OPr81cnGH6//L0npjnzyTgajcA6+jz//T3bRp8Obfsa/csVMz6NXly73DN7V3bNu7RrTLRlOx+4bGRljQRDEYQjIHNc1DOH7TVIKGJbL5WYzyJIYuSiVikop32+S1lzwQqHYaDQcx3Zdz/cbWmvGWBRFoMmy7Vw+BwBhGBHpXC5XrdY4YwSUy+UMw6jVaoxxrVVrcGatVi8WC2mahc0mMFYo5DnnjUYjn88rqZpBs1gsBkFoGEaaJo5j247b9JtSynw+t2s1TYgiS4YTf3Dc+AnV4Uo9TCaNnyzy6dm1D31jv0uu67/7okXvcIWrdwBH7Cpn9MifxYcCOPYwbVw5nM9NZMLIsoSAbNtJ04DbRYTdhXBaa8/zcvlcK2tDRB0dba3rV0oVCnnGCkSglBJCtHe0tYjZKIpKxQIyFoWB6zpSSiEM13W0JsZYkiYIwBlyYWUyy+U8AFJSAVEUhblcjjGWpokQQmZZqVRgjNm25XkOEbQG8pbLZaWUYRodTruUqlgsEGnXdbTWMstc12ld+W4MIZHSaWSYFmMsjhNgor29MNjff2R5fkr4dOMxTRqB7WjGdlZ7p0kY+OSL2t9gjJ2qpo2t9fqQc52gnkRhkCvkB/v7QaZoiN3bZ623IWalVFv/afPXISIlFSKmabLfjJnHn3SaZdsvPPvMc39++sKPXvSLq6+cPHnKnEMP+9+bb7rwI5/wcrksTe+49eaOjq63n3JGJrPVK1c8cM9v//mDFz/28INxHL/z3PMu/+F/H3DwoYccdvj11/wcEbf+3JbmJaLWwU2KWG254NdwOBipKEuDcrktiWOlyXYcMGRb1HPghIOGkuzp+qOM4U4bLNhOE/u2BWs26CdfMqiojp+X1Op+d1e7Bt70m47jMIYqDWifNnwjYpKkx7/9FGR48/W/WPDS867rTp42QzCWL5YnTJyCCNNnzfz9vXclUXTWu97T3t5pWuavb7zu+JNOnT5zds/osZ1dPSuXLQGCD3zo46ef9c5XF7ySxDFj+3jpJJ2FADqXy/m+rwE729v8qHGIeWSPbS6MFiwJF9nM0bRnQLdKHZWi3z7MMGFnHweoBkuFMhhGkiVaKdfLZWmTVAq4z7AmIsuy77/7Tgb0zx/6yPRZB5DScRQmaRLHYZomCND0m2899vjR48YvWbQgTiIvVzjltDMH+nv7+3tVJuMolGn2i5//bObs/VevXPHYww+5nvfaQrpX+Q3KsqRp2x4ixnGsuejqKDeH4lPa34kaHq7dG8qAo9iLriylwTDo3kepvkZM3J8O3q8aBtRW9Ah4o14vFotaSZ0F+xZozjkR3nzDdQtffvm0M89Ks9Sy7ImTpoyfOEkppaRyPW/tmlXXX3Pl/ff8tlgsSZk++afHL//B9wb6+ry819kzar9Zs7Ms7evrXb9+LeC+bf8nRKazSGZxsVRqiXPOyZGRjAknzSsf0ZumD1TuaCUSd15Lv6tecNPE/iE6ZBafdVhmxXTbQ+6Y0R29AxWVxsVSKUmTLIu56eGuSwtxb24DYyxN04MPnXvyqWeW29oee+ShRa+8TATHnfj2YrF49x231eu1iZMnP/LgA+vWrBaG0dnVk6TpA/f8FgAM0+zpGT1u3IT9DzpkoG+jZduDg/0bN6w3DHM317P51x1f7MpByKIKY6yt3FYZHEyRTZ0wqZ5UzuUfenvHEQ/VH/vZ+v+2mb2r+qtdzr3jDJoBnXKMdfcvgjTEo89zeG7OK4sXqjAoFgqWbfX2bnRyXdwpwc4eT6KR3oTW1dMIJwmbX299pPUrACRxbFoWIiZJYllmHMW26yoplZSmZWVZJoRgjBNpqRQQmKZBAFopKeUm35S11ltgW6G2+f9vQZOAYOQyNmf3RxjXncLNmE4aUaOvo7MLESuVijadw+fM7l9a+c3UhyfZ3Rcvu+hXvdfkjcKumoXErsvrwXHwoaez558zDj0+fv+p4Q9/05g+adyLi5b5jXqhOM6y7Sypc9PbMUokItM029vb4ziO4xgAbMdBgCAMbMt2XTdJEqVUlmUAIKVMksSyrHK5HCeJEEJJmaSJaZhtbTwIglxbm23bYRhKKU3TbAYBAti2bRhGtVpVSnV3dwOA7/tccMu0sixjjLmuG8exECIMQkXKNEzLssIw1FpLKT3PcxwniqIsy3K5XLVa9TwPEWu1mlJqJ1hrJeO6MAzP8zZu3KCATxkzfjgenG+cMd3tfqG56v7KnTa3dzMAme0mpBcckph+eosBEf7zOegZazo7xjiOrRFrtVpHe4fKEpXUd7wqIjItM4wiKWV7e3u5rU2RTrLUy+cJIEuzOI5zuVw+n7dtGxFbZLfl2Ll8LpfP5Qp5wzRt18kXC+W2chzHrZvheG6ukC+WS6VyiTEWx7Fpmq0ortFodHR2Wo5tWiZjLAzD1o0sFouO61i27eY927bz+bxpmqZhEFCWpi2Uh4aG2jrbHc8tFAo791KQqcRP06itrb3ZDJQiENbEST3Jev2BMZ/UBL8ZuraaVgy2u8YDtvsqeduhWx+glQus4kR14VmNJcs3TJkwWoLwG3XOhevl0qiuZQQ7aGrS1BK6LMuklAbnOS9nCkMrlcnMcRxE5Jy7rut5nmEYRJQmaX9ffxInURgaQpCmWrWW83JaaaWUUoohq9VqtmUBASJ6nmdZVivkcRwnyzKtdKPeKBaLjDEpJRH5vq+15shIUeuIYRi5fD6JE611ayWBXC6npcrSNAgC0zS3DwsRSSVZXLUs23GcWrWimJg4ZvTagdXHsTMOLE5YHg/+b/91Jjf1bkPl1xiMYgisNzShccr8ZPYk/sv/bZRKUyuNqlYqS6P2js5GvYagmels3UyAiFJKzliSJL7vR2HY+lZ+oyGlzLJMKVWr1ZrNZusxV0pJKcMg0ErFURSFkVY6iqIkjoMgaImzlDJN0yROgmbQOqi19n1/cyRSq9WUlIZhNBqNJElaN7j1EVmaxXFERLVaTUqplAqCoCXyaZq2wvc4ioMgaIUtuI1+hywelknUM2pUrVbLpELDmrHfhIElw/819WfdvPjTvv+5Z+BWV3i7nzf4GkBrAmHAS6/CO44UYw9IcxpueYDvN2nU2r4hyhLLtm3bbjbqjHMm3O00dUurtvJ9aZqmaUoji6bozbYrTdMoilpC1ErutTzfTEqtVEuD601xZOu1Uqr1czOfGcdxmqYtFdS6Ja3XrXepTVsURS0JiOO4dXdbp7WObz55W5RRpn4SVIqlsmEY1eFKhsbsqdMGmhvezT5ybs/pi6INn1/24RQS9lrh22sHTgaHZqD/8woBTfb+c2Dm+LVS2T1dZRL20MBAzstZtp2FNZ2F2xWOjLD1iC3vjTGGW22bvbrWyimbGdzWa7bpnK3Px203salXl3M+khZorT206fV272qd31JZmz9o6+PbengEyEgmWVQRwiiVSoP9AyDsQr6Yb7OcjeWPTPh0puHK3u/1Jht349XtBdBSgZej39yvf/97lxeT714CK1ctnjxuvGacCWNgcKCrqwuIsnhI622qO2jPJobR6x0stvmNb/CDdnGcgVZpNKSk6u7uGR4e1giScMaUCYsWL75k3DdGG96TzRdu7L3SFe6eTJTeo0mODJlS6oUl/Py3w8RZWVJJf/+UO3VS5/r+YZAJ5zxfLDZqVUaKG97f/dC7TVpDRpU09js6uoioXqtmaMyYPNmXlWOi0/518r/WVPq5ZR9cES61ub0n02D3CGgisEy2dgMJZs4/OnnLHHHHfTXFJgqe+aGMw0YuXzCEETQbHImb7t871Igo42oaVnP5opfPDfT1gmGXC6WxEzqaS9KfzrqpxNyr+q68ZsP/5I38Hg5I39PZpERgmvDYszD/QHvS/sHRs40fXTu8336z+yuDABj49fb2DqUpDhsMNTdyf19rSW+LMpNxLQuHTdvpaG/v692IwiQUh86Z9crLC3806cZDczP+HCz81JL37RXHsBfTdjmDLNNPvsjeewIbOyMZV8Arbw4OO3j6qo0DBvKm3+jq6kzSLIl9ZMDF3yfWyFRST6MKF0ZPT09/fz8hppodcdCBC9e88nH3i+ePfc+gDD++5L0roiX23oyQ3puxxgSWiev71Lo+413Hq/0PUtWB+KGnrANnjV+5fsDgFAZ+Z2dXkmRp3ETU3Pg70yGIqOJaGlUY4z09PUNDFaVkrOGQ2bP74965jeO/MeO7pOE/1nz+zv5fF4zCXq2qsHcT0TWBY+Ozr+i8ZR85Lz3hSPbU09Xla9unTe5Y01sRgGHY7OzqSpM0jXwEzYQDyN70CLfmMqGMq1k43JLlSmU4S5OY2KypM7Qd5Vd2X3PgLTkwrhu84ZurLs0Zub1dQ2uvZ/wTgGnS/Y/TwZPN6QdEpx8r7v7dUG+l3NPlDtRCBjoKm13d3VKqKKojSS6c3Q+HeBNIMgfSMqqkYc207ZYsZ2mSIp8werTIZWq5ecOBd43mpT80//SJJe8RwHHvR+q/nuVBGINMgm2y318Lc48MKuud4y8gP51qOcni5RtdDkSyZ9QYv+HXasOm6Qi3nQuXSL8JJ9EjIskkjSpZEuTyhXK53N/fr7WONU4cO8p2XX8h/GbebbOcKS9Fy979yvGDaZ/FLb33i3C8rk4/DaagIFZnfQyXveK2j40e/CUrGCvi0D70wP1CDciM3vXrPdft6RklZZI2+7KkBoj4plIjiIAg00YcbMzSsL2jq1gq9W7coIlijbOnTzZFexK8fNfXjprOp6xM+i5c9I6+ZL3Nbf26ljp5nYuSEaFpYK1B9z/OTztSjJkS/tMJxv1/GBqsd8+a2rOmvyK4aNSHLdtub+8IwyiN6qhTZAZyY2SY/t9WtJGRymRUScIKZ2LUqNFENNDXi8JMNDv0wP2yzDXUK3dfYc0864m+ZuO8W362KHk6Jwrq9S6594YWJRMcGk2YPkHc9XM1bXYQVd3z/kU9t2zqtCkdz7z0SpZKJlPbsTs7O/2GP1ytIOemXRBWEVojmf7acBMAAmOglUoaWVxXMi2WyqVSeXi4EjR9EA4wnHvQjA192ejiglt/ZnSNDfpWOe/8bPTks1hwXPkGUr1vaJk9TWA70D+kbr2fHzPHmrBf+J5TxPq1w/c9ns09eFacxLUwRlKNWq1QLJTKbVmcRmGDZIwIjJnIxF9XG3MArdMgCwfTqG6YRk/PaMMw+ns3pjLL0CyXcocdtP+CV3uPP3j5b37Ki13B8oW5ky/kL7xiFDxDvrGE+htdz5A02BZWfX3LPWz2RGu/WdFJx7GefPjrO4fGjduvq93rqzQQWeA3pJTtnR2u60ZRmEZ1UhEAITOQ8U2LmO5zAafWEqmADEDqrJlFlSSsMsSOzq5yuVyv1YYrwyQMCXzm1PGjR41fuHDR5y+sfPerJHLpE4/nz/wIrVgvCzmQb3gh5X2wQqcmsAyMU33z3Zg3rXlz0jlz5Ylz8a4HNgzVyzOnjQnipJlkpGSjXhVCdHR0WpadJFEcNigLCLIWXQpMwJbmSXzjWhhHJlanOmmkYSWJaoyxtvb29o6OJEkH+vsyqTJm5vO5g2ZN833Imguuuyw577wUSF//K++8z0KtoTwX1T5YCXWfLIU6UuyLjMPdf9Br1ltvm4MTZibvO4UNDQ3/4clk8qT9ujqKlXpTKkjjKPB907La29td15NSxqEvE1/LCEkhADBgyLG1FADuIJ6v4UWwTTS0BpXKrCmjahoOyyywTKuzs6tcbpNSDg4OhGGkmIGmMWvapHGjJixfvvrkt6y58fswY04YV8wvfNu95DLJhbZMVPuoAmdfrnQPCJyB78OcGeInX1dHHhUDY489bH7l/9LawSnjx3UMDg+sWteXJYkAxZDyuXy+WEDAIAh8v5GmCQDjwhDCYcICbiATI8tVj/x7AEDaPJmFWj2hSEBAGkjSyOLJsZKJkimQNkwrl8vnch4i8/1G0GhIrSVyZlgTxnaP7hy1fmO17Kz46r+oM05NgasFL3if/Bp79M9ZLoetnv99hs2+BHoksQJ+QJbBvvIJ8ZkPxm5XKofNn91EV9zsST5l4oSeDX3rVqzZkCQJ15KRtm07XyjYtq01xXEUBkGcJEpKQmDIGDcZF8g4Q97KoGzp29UaSBMprbVWmVYZadlKuFi27Tiu67qMsSSJ/YYfx7ECVMgM05o4fvTEMePXb6gmzeX/fHbjMx9Eb1Siho0rfuV8+Xuq6qtCHqXc17Z4nwPd4vmUhjCkeQcZ37qU5h8dgkuDK8zLb4Tf/K5IfGK57ISxv2ZDb6PRBJUxIsHQskzXy9mOwxlXSmVZmiRJmiYyk0pJrTVtmwxppZ44Y4xzQxiGaVqWZZimEEIr1UrsxnEstSZgxEUu544f01PItdVqqYxXn3HM0CcvwPGzMkjp2We8L/43/v4JaTtkiH2jlP8aQLdQ4Bz8gARj57+TX/phNX3/CATrW8qv/LW6/cFiPR7T3lYSpqxUh/sHq0EYgpKMNEcQgpumaduOMIRhWIyxFrVAQFrpzaVPjLGReXRERCN1GmmSJEmSSakINDLgwnHszo5yV1u7VsZwxbf4htOOqX/sPJq0vwKgdUvtH1wrrviVDiJVyKPSf6kle/5SQG9mRYggCKitaFz8Xvzoe9Lx02MwobHauPUBuPlea+nadtvpKRZt5LJaqw7X6o0gkGlGSrbWOxaIyJAxxhlrpV5x0yC4Vm251lppTVqr1hiVVvLVMPOe21YqtJXKSGbdj6Pm0PiegX86KXr3qdg1lUCrgZXmNbca//NL2DggXRc4/4sI8l8J6M0BZJpBHFNPFzvvdOOif0pnH5hCTkONvfACu/33+Oiz9sZKGxPt+XzOtFCpJIyCRrPZDOMkSWSmlJIjxXxbRw3IEJExZJxzwR3bchyrkMt7tmMYbpJS0IxkWuksVd46Jzz7RH3kXIA2BRGueNX65e3GDXfQqvXStMg2/4KC/FcFGjYNaUhSShIo5Pipx7ILztbzj8jsURlwhCF8eRE88jQ+9Yq5bI1XDfKMFw3TtCzTNIXWKZHSWmlSSinYpDoE54xxZByIcWbIjJIky7JEynrRaU4ZG8w9IJl/uD7sAMAuAlCq33j8OfP6O8SdD+lKVZmmti32V4D4rwr01nBnkqIIEPkB0/HM4/DUt6m5+yeik8ACiCDppUUr8OWlbMlqvnoj7x+y6oERJXamOJCQJNgI0MBQMsg4l7aVFbykqy2bOFpOn6gOmKZmTcXcKACPQAIN8RdeFfc9xu98EJ9fqJQi2ybTwL8axH8DoDfbScaBCKKIlELDZPtPxSMPxbcdrufuLyeMlVBWYDHQADFCE6IGDdew5qsg4tUGIhIAkMZCQRccnc+L9qL2igA5AIeAIaQENb6xVzy3kD/yDP7xWXhlKUWRRkaugy2P6K+/SN3fAOitTSUy0IrCGEkBAHa04bSJeOB+OHsaTZ9Ek8bI7jYqFDU4BKYGDiC2Sm0oAAWQMIix2eD9w7imF5euZAtX4MuLcelq3TdIrXlbrg2CgybQf7sFLf+WQG+NOEMgglRSkrRGCCEwLHpQLmFHCdpLWC6Q50B7kWsiAGAMaw3VCKnu80pNV6pYqUOjqbQaWVHKNMk0kCFo2n7u0/+7QG9DV+DIiFNNoBRJCZlqde3tdFgXtmRW8NY+MtWHaGT9iTfRV3tTAb0j7ptSTlsPnt+GZ2qNPdo0UPlN+1VAwJt4e/PDtxfqEf6x/QPofwD9j22vt/8PjG1eEuLLfWoAAAAASUVORK5CYII=";
  const w=window.open("","_blank","width=900,height=700");
  w.document.write(`<!DOCTYPE html><html><head><title>${title}</title><style>
    *{margin:0;padding:0;box-sizing:border-box;font-family:Arial,sans-serif}
    body{padding:32px;color:#0F172A;font-size:13px}
    .hdr{display:flex;align-items:center;gap:16px;margin-bottom:6px;padding-bottom:14px;border-bottom:3px solid transparent;border-image:linear-gradient(90deg,#00D4FF,#FF00FF,#FFD700) 1}
    .hdr img{width:56px;height:56px;border-radius:10px}
    .hdr-txt .gks{font-size:24px;font-weight:900;letter-spacing:-1px;line-height:1}
    .hdr-txt .gks span:nth-child(1){color:#00D4FF}
    .hdr-txt .gks span:nth-child(2){color:#FF00FF}
    .hdr-txt .gks span:nth-child(3){color:#FFD700}
    .hdr-txt h1{font-size:16px;font-weight:700;color:#0F172A;margin:2px 0 1px}
    .hdr-txt .sub{font-size:11px;color:#64748B}
    .meta{font-size:11px;color:#64748B;margin:10px 0 20px;padding-bottom:10px;border-bottom:1px solid #E2E8F0}
    h2{font-size:14px;font-weight:700;margin:24px 0 10px;border-bottom:2px solid #E2E8F0;padding-bottom:6px}
    table{width:100%;border-collapse:collapse;margin-bottom:20px;font-size:12px}
    th{background:#F8FAFC;padding:8px 10px;text-align:left;font-weight:600;color:#64748B;border-bottom:2px solid #E2E8F0}
    td{padding:8px 10px;border-bottom:1px solid #E2E8F0}
    .g{background:#D1FAE5;color:#065F46;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
    .r{background:#FEE2E2;color:#991B1B;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
    .b{background:#EBF2FF;color:#1E40AF;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
    .sum{display:flex;gap:32px;background:#F8FAFC;padding:14px 18px;border-radius:8px;margin-bottom:20px}
    .sum label{font-size:11px;color:#64748B;display:block;margin-bottom:3px}
    .sum span{font-weight:700;font-size:15px}
    .ftr{margin-top:40px;padding-top:12px;border-top:1px solid #E2E8F0;display:flex;align-items:center;justify-content:center;gap:10px;font-size:11px;color:#94A3B8}
    .ftr img{width:22px;height:22px;border-radius:4px;opacity:0.6}
    @media print{button{display:none!important}}
  </style></head><body>
  <div class="hdr">
    <img src="${LOGO}" alt="GKS"/>
    <div class="hdr-txt">
      <div class="gks"><span>G</span><span>K</span><span>S</span></div>
      <h1>Global Karunia Supply</h1>
      <div class="sub">Digital Printing Distributor &nbsp;·&nbsp; ${title}</div>
    </div>
  </div>
  <div class="meta">Diekspor: ${new Date().toLocaleString("id-ID")} &nbsp;|&nbsp; GKS ERP System</div>
  ${html}
  <div class="ftr"><img src="${LOGO}" alt="GKS"/> Global Karunia Supply — Digital Printing Distributor — Dokumen digenerate oleh GKS ERP</div>
  <script>setTimeout(()=>window.print(),400)</script>
  </body></html>`);
  w.document.close();
}
const buildInvExport = ps=>({
  cols:["SKU","Nama","Kategori","Gudang","Stok","Min Stok","Harga (IDR)","Nilai Stok (IDR)","Status"],
  rows:ps.map(p=>[p.sku,p.name,p.category,p.warehouse,p.stock,p.minStock,fmtR(p.price),fmtR(p.stock*p.price),p.stock<=p.minStock?"Kritis":"Aman"]),
  html:`<h2>Inventori</h2><table><thead><tr>${["SKU","Nama","Kategori","Gudang","Stok","Min","Harga","Nilai","Status"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${ps.map(p=>`<tr><td>${p.sku}</td><td><b>${p.name}</b></td><td>${p.category}</td><td>${p.warehouse}</td><td><b style="color:${p.stock<=p.minStock?"#EF4444":"#10B97B"}">${p.stock}</b></td><td>${p.minStock}</td><td>${fmt(p.price)}</td><td>${fmt(p.stock*p.price)}</td><td><span class="${p.stock<=p.minStock?"r":"g"}">${p.stock<=p.minStock?"Kritis":"Aman"}</span></td></tr>`).join("")}</tbody></table><div class="sum"><div class="sum-item"><label>Total SKU</label><span>${ps.length}</span></div><div><label>Total Stok</label><span>${ps.reduce((s,p)=>s+p.stock,0)} unit</span></div><div><label>Nilai</label><span>${fmt(ps.reduce((s,p)=>s+p.stock*p.price,0))}</span></div></div>`
});
const buildMovExport = ms=>({
  cols:["Tanggal","Jenis","Tipe","Produk","Qty","Ref","Catatan","Sebelum","Sesudah"],
  rows:ms.map(m=>[m.date,m.type==="in"?"Masuk":"Keluar",m.subtype||"-",m.product,m.type==="in"?m.qty:-m.qty,m.ref,m.note,m.before??"-",m.after??"-"]),
  html:`<h2>Mutasi Stok</h2><table><thead><tr>${["Tgl","Jenis","Tipe","Produk","Qty","Ref","Catatan","Sebelum","Sesudah"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${ms.map(m=>`<tr><td>${m.date}</td><td><span class="${m.type==="in"?"g":"r"}">${m.type==="in"?"Masuk":"Keluar"}</span></td><td>${m.subtype||"-"}</td><td><b>${m.product}</b></td><td><b style="color:${m.type==="in"?"#10B97B":"#EF4444"}">${m.type==="in"?"+":"-"}${m.qty}</b></td><td>${m.ref}</td><td style="color:#64748B">${m.note}</td><td>${m.before??"-"}</td><td>${m.after??"-"}</td></tr>`).join("")}</tbody></table>`
});
const buildOrdExport = os=>({
  cols:["No. Pesanan","Tanggal","Pelanggan","Total (IDR)","Status","Invoice"],
  rows:os.map(o=>[o.id,o.date,o.customer,fmtR(o.total),o.status,o.invoiceId||"-"]),
  html:`<h2>Pesanan</h2><table><thead><tr>${["No","Tgl","Pelanggan","Total","Status","Invoice"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${os.map(o=>`<tr><td><b style="color:#1B5FBF">${o.id}</b></td><td>${o.date}</td><td><b>${o.customer}</b></td><td><b>${fmt(o.total)}</b></td><td><span class="${o.status==="delivered"?"g":o.status==="cancelled"?"r":"b"}">${o.status}</span></td><td>${o.invoiceId||"—"}</td></tr>`).join("")}</tbody></table><div class="sum"><div><label>Total</label><span>${os.length}</span></div><div><label>Nilai</label><span>${fmt(os.reduce((s,o)=>s+o.total,0))}</span></div></div>`
});
const buildInvoExport = ivs=>({
  cols:["No Invoice","No Pesanan","Pelanggan","Tgl","Jatuh Tempo","Subtotal","PPN","Total","Status","Tgl Bayar"],
  rows:ivs.map(i=>[i.id,i.orderId,i.customer,i.date,i.dueDate,fmtR(i.total),fmtR(i.tax),fmtR(i.grandTotal),i.status==="paid"?"Lunas":"Belum Bayar",i.paidDate||"-"]),
  html:`<h2>Invoice</h2><table><thead><tr>${["No","Pesanan","Pelanggan","Tgl","Jatuh Tempo","Subtotal","PPN","Total","Status","Bayar"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${ivs.map(i=>`<tr><td><b style="color:#1B5FBF">${i.id}</b></td><td>${i.orderId}</td><td><b>${i.customer}</b></td><td>${i.date}</td><td>${i.dueDate}</td><td>${fmt(i.total)}</td><td>${fmt(i.tax)}</td><td><b>${fmt(i.grandTotal)}</b></td><td><span class="${i.status==="paid"?"g":"r"}">${i.status==="paid"?"Lunas":"Belum Bayar"}</span></td><td>${i.paidDate||"—"}</td></tr>`).join("")}</tbody></table><div class="sum"><div><label>Tertagih</label><span>${fmt(ivs.reduce((s,i)=>s+i.grandTotal,0))}</span></div><div><label>Lunas</label><span style="color:#10B97B">${fmt(ivs.filter(i=>i.status==="paid").reduce((s,i)=>s+i.grandTotal,0))}</span></div></div>`
});
const buildCFExport = es=>({
  cols:["Tanggal","Keterangan","Kategori","Ref","Sumber","Jenis","Jumlah (IDR)"],
  rows:es.map(e=>[e.date,e.description,e.category,e.ref||"-",e.source==="invoice"?"Invoice":"Manual",e.type==="in"?"Pemasukan":"Pengeluaran",e.type==="in"?fmtR(e.amount):-fmtR(e.amount)]),
  html:`<h2>Arus Kas</h2><table><thead><tr>${["Tgl","Keterangan","Kategori","Ref","Sumber","Jenis","Jumlah"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${es.map(e=>`<tr><td>${e.date}</td><td><b>${e.description}</b></td><td>${e.category}</td><td>${e.ref||"-"}</td><td>${e.source==="invoice"?"🧾":"✏️"}</td><td><span class="${e.type==="in"?"g":"r"}">${e.type==="in"?"Pemasukan":"Pengeluaran"}</span></td><td><b style="color:${e.type==="in"?"#10B97B":"#EF4444"}">${e.type==="in"?"+":"-"}${fmt(e.amount)}</b></td></tr>`).join("")}</tbody></table>`
});
const buildCustExport = (cs,os)=>({
  cols:["Perusahaan","PIC","Telepon","Email","Alamat","Pesanan","Total (IDR)"],
  rows:cs.map(c=>{const co=os.filter(o=>o.customer===c.name);return[c.name,c.contact,c.phone,c.email,c.address,co.length,fmtR(co.reduce((s,o)=>s+o.total,0))];}),
  html:`<h2>Pelanggan</h2><table><thead><tr>${["Perusahaan","PIC","Telepon","Email","Alamat","Pesanan","Total"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${cs.map(c=>{const co=os.filter(o=>o.customer===c.name);return`<tr><td><b>${c.name}</b></td><td>${c.contact}</td><td>${c.phone}</td><td>${c.email}</td><td style="color:#64748B">${c.address}</td><td>${co.length}</td><td><b>${fmt(co.reduce((s,o)=>s+o.total,0))}</b></td></tr>`;}).join("")}</tbody></table>`
});

const buildPOExport = (pos)=>({
  cols:["No PO","Tanggal","Supplier","Status","Subtotal","PPN","Total","Tgl Terima","Catatan"],
  rows:pos.map(p=>[p.id,p.date,p.supplierName,p.status,fmtR(p.subtotal),fmtR(p.tax),fmtR(p.total),p.receivedDate||"-",p.notes||""]),
  html:`<h2>Purchase Order</h2><table><thead><tr>${["No PO","Tanggal","Supplier","Status","Subtotal","PPN","Total","Tgl Terima"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${pos.map(p=>`<tr><td><b style="color:#00B4D8">${p.id}</b></td><td>${p.date}</td><td><b>${p.supplierName}</b></td><td><span class="${p.status==="diterima"?"g":p.status==="dibatalkan"?"r":"b"}">${p.status}</span></td><td>${fmt(p.subtotal)}</td><td>${fmt(p.tax)}</td><td><b>${fmt(p.total)}</b></td><td>${p.receivedDate||"—"}</td></tr>`).join("")}</tbody></table><div class="sum"><div><label>Total PO</label><span>${pos.length}</span></div><div><label>Total Nilai</label><span>${fmt(pos.reduce((s,p)=>s+p.total,0))}</span></div></div>`
});
const buildReturnExport = (rets)=>({
  cols:["No Retur","Tanggal","Tipe","Referensi","Pihak","Total","Alasan","Status"],
  rows:rets.map(r=>[r.id,r.date,r.type,r.refId,r.customer||r.supplier||"-",fmtR(r.total),r.reason,r.status]),
  html:`<h2>Retur Barang</h2><table><thead><tr>${["No Retur","Tanggal","Tipe","Ref","Pihak","Total","Alasan","Status"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${rets.map(r=>`<tr><td><b style="color:#00B4D8">${r.id}</b></td><td>${r.date}</td><td>${r.type}</td><td>${r.refId}</td><td>${r.customer||r.supplier||"-"}</td><td><b>${fmt(r.total)}</b></td><td style="color:#64748B">${r.reason}</td><td><span class="${r.status==="disetujui"?"g":r.status==="ditolak"?"r":"b"}">${r.status}</span></td></tr>`).join("")}</tbody></table>`
});
const buildDiscountExport = (discs,customers,products)=>({
  cols:["Nama","Tipe","Nilai","Min Order","Pelanggan","Produk","Status","Keterangan"],
  rows:discs.map(d=>[d.name,d.type==="persen"?`${d.value}%`:`Rp${d.value.toLocaleString()}`,d.value,fmtR(d.minOrder),d.customerId?customers.find(c=>c.id===d.customerId)?.name||"-":"Semua",d.productId?products.find(p=>p.id===d.productId)?.name||"-":"Semua",d.active?"Aktif":"Non-Aktif",d.desc]),
  html:`<h2>Daftar Diskon & Harga Khusus</h2><table><thead><tr>${["Nama","Tipe","Nilai","Min Order","Pelanggan","Produk","Status"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${discs.map(d=>`<tr><td><b>${d.name}</b></td><td><span class="b">${d.type==="persen"?"Persen":"Nominal"}</span></td><td>${d.type==="persen"?`${d.value}%`:fmt(d.value)}</td><td>${d.minOrder>0?fmt(d.minOrder):"—"}</td><td>${d.customerId?d.customerId:"Semua"}</td><td>${d.productId?d.productId:"Semua"}</td><td><span class="${d.active?"g":"r"}">${d.active?"Aktif":"Non-Aktif"}</span></td></tr>`).join("")}</tbody></table>`
});

function ExportBtn({onExcel,onPdf}){
  const [open,setOpen]=useState(false);
  return(
    <div style={{position:"relative",display:"inline-block"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{background:"#FFF",color:C.text,border:`1px solid ${C.border}`,borderRadius:8,padding:"7px 14px",cursor:"pointer",fontWeight:600,fontSize:13,display:"flex",alignItems:"center",gap:6}}>
        ⬇ Export <span style={{fontSize:10,opacity:0.6}}>▼</span>
      </button>
      {open&&(<>
        <div onClick={()=>setOpen(false)} style={{position:"fixed",inset:0,zIndex:99}}/>
        <div style={{position:"absolute",right:0,top:"calc(100% + 4px)",background:"#FFF",border:`1px solid ${C.border}`,borderRadius:10,boxShadow:"0 8px 24px rgba(0,0,0,0.12)",zIndex:100,minWidth:160,overflow:"hidden"}}>
          <button onClick={()=>{onExcel();setOpen(false);}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 16px",border:"none",background:"transparent",cursor:"pointer",fontSize:13,fontWeight:500}}>📊 Excel (.xlsx)</button>
          <div style={{height:1,background:C.border}}/>
          <button onClick={()=>{onPdf();setOpen(false);}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 16px",border:"none",background:"transparent",cursor:"pointer",fontSize:13,fontWeight:500}}>📄 PDF / Cetak</button>
        </div>
      </>)}
    </div>
  );
}

// ─── INVENTORY PAGE ───────────────────────────────────────────────────────────
function InventoryPage({products,setProducts,movements,setMovements}){
  const [tab,setTab]=useState("daftar");
  const [alert,setAlert]=useState(null);
  const [editProd,setEditProd]=useState(null);      // product being edited
  const [confirmDel,setConfirmDel]=useState(null);  // product id to delete
  const [editMov,setEditMov]=useState(null);        // movement being edited
  const [confirmDelMov,setConfirmDelMov]=useState(null);

  // stok masuk/keluar forms
  const [fMasuk,setFMasuk]=useState({productId:"",qty:"",ref:"",note:"",date:today(),supplier:""});
  const [fKeluar,setFKeluar]=useState({productId:"",qty:"",ref:"",note:"",date:today(),reason:"penjualan"});
  // opname
  const [opRows,setOpRows]=useState([]);
  const [opDate,setOpDate]=useState(today());
  const [opDone,setOpDone]=useState(false);

  const lowStock=products.filter(p=>p.stock<=p.minStock);

  // ── product CRUD ──
  const saveProduct=(p)=>{
    if(editProd.id){setProducts(prev=>prev.map(x=>x.id===p.id?p:x));}
    else{setProducts(prev=>[...prev,{...p,id:uid()}]);}
    setEditProd(null);
  };
  const deleteProduct=(id)=>{
    setProducts(prev=>prev.filter(p=>p.id!==id));
    setMovements(prev=>prev.filter(m=>m.productId!==id));
    setConfirmDel(null);
  };

  // ── movement CRUD ──
  const saveMovement=(m)=>{setMovements(prev=>prev.map(x=>x.id===m.id?m:x));setEditMov(null);};
  const deleteMovement=(id)=>{setMovements(prev=>prev.filter(m=>m.id!==id));setConfirmDelMov(null);};

  // ── stok masuk ──
  const handleMasuk=()=>{
    const prod=products.find(p=>p.id===Number(fMasuk.productId));
    if(!prod||!fMasuk.qty||Number(fMasuk.qty)<=0)return;
    const qty=Number(fMasuk.qty),before=prod.stock,after=before+qty;
    setProducts(prev=>prev.map(p=>p.id===prod.id?{...p,stock:after}:p));
    setMovements(prev=>[...prev,{id:uid(),date:fMasuk.date,type:"in",subtype:"pembelian",product:prod.name,productId:prod.id,qty,ref:fMasuk.ref||`PO-${Date.now()}`,note:fMasuk.note||`Masuk dari ${fMasuk.supplier||"supplier"}`,before,after}]);
    setAlert({ok:true,msg:`✅ Stok ${prod.name} bertambah ${qty} (${before}→${after})`});
    setFMasuk({productId:"",qty:"",ref:"",note:"",date:today(),supplier:""});
  };

  // ── stok keluar ──
  const handleKeluar=()=>{
    const prod=products.find(p=>p.id===Number(fKeluar.productId));
    if(!prod||!fKeluar.qty||Number(fKeluar.qty)<=0)return;
    const qty=Number(fKeluar.qty);
    if(qty>prod.stock){setAlert({ok:false,msg:`⚠️ Stok tidak cukup! Stok saat ini: ${prod.stock}`});return;}
    const before=prod.stock,after=before-qty;
    setProducts(prev=>prev.map(p=>p.id===prod.id?{...p,stock:after}:p));
    setMovements(prev=>[...prev,{id:uid(),date:fKeluar.date,type:"out",subtype:fKeluar.reason,product:prod.name,productId:prod.id,qty,ref:fKeluar.ref||`OUT-${Date.now()}`,note:fKeluar.note||`Keluar – ${fKeluar.reason}`,before,after}]);
    setAlert({ok:true,msg:`✅ Stok ${prod.name} berkurang ${qty} (${before}→${after})`});
    setFKeluar({productId:"",qty:"",ref:"",note:"",date:today(),reason:"penjualan"});
  };

  // ── opname ──
  const startOpname=()=>{setOpRows(products.map(p=>({...p,aktual:p.stock,note:""})));setOpDone(false);};
  const submitOpname=()=>{
    const ref=`OPN-${Date.now()}`;
    const moves=[];
    opRows.forEach(r=>{
      if(r.aktual===""||r.aktual===r.stock)return;
      const diff=r.aktual-r.stock;
      moves.push({id:uid(),date:opDate,type:diff>0?"in":"out",subtype:"opname",product:r.name,productId:r.id,qty:Math.abs(diff),ref,note:r.note||`Opname (${diff>0?"+":""}${diff})`,before:r.stock,after:r.aktual});
    });
    setProducts(prev=>prev.map(p=>{const r=opRows.find(x=>x.id===p.id);return r&&r.aktual!==""?{...p,stock:r.aktual}:p;}));
    setMovements(prev=>[...prev,...moves]);
    setOpDone(true);
  };

  const tabs=[{id:"daftar",l:"📋 Daftar Produk"},{id:"masuk",l:"📥 Stok Masuk"},{id:"keluar",l:"📤 Stok Keluar"},{id:"opname",l:"🔍 Stok Opname"}];

  return(
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total SKU",v:products.length,c:C.primary,i:"📦"},
          {l:"Total Stok",v:`${fmtN(products.reduce((s,p)=>s+p.stock,0))} unit`,c:C.success,i:"📊"},
          {l:"Nilai Inventori",v:fmt(products.reduce((s,p)=>s+p.stock*p.price,0)),c:C.warning,i:"💰"},
          {l:"Stok Kritis",v:lowStock.length,c:C.danger,i:"⚠️"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div></div>
              <span style={{fontSize:22}}>{k.i}</span>
            </div>
          </div>
        ))}
      </div>

      {/* Tab bar */}
      <div style={{display:"flex",gap:4,marginBottom:16,background:"#FFF",borderRadius:12,padding:6,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",border:`1px solid ${C.border}`,width:"fit-content"}}>
        {tabs.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?C.primary:"transparent",color:tab===t.id?"#FFF":C.muted}}>{t.l}</button>
        ))}
      </div>

      {/* ── DAFTAR PRODUK ── */}
      {tab==="daftar"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div style={{fontWeight:700,fontSize:15}}>Daftar Produk</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setEditProd({id:null,sku:"",name:"",category:"",stock:0,minStock:0,price:0,unit:"Unit",warehouse:"Gudang A"})} style={btn(C.primary)}>+ Tambah Produk</button>
              <ExportBtn onExcel={()=>{const e=buildInvExport(products);exportExcel([{name:"Inventori",...e}]);}} onPdf={()=>{const e=buildInvExport(products);exportPDF("Inventori",e.html);}}/>
            </div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["SKU","Nama Produk","Kategori","Gudang","Stok","Min","Harga","Nilai Stok","Status","Aksi"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
            <tbody>
              {products.map(p=>(
                <tr key={p.id} style={{borderBottom:`1px solid ${C.border}`,background:p.stock<=p.minStock?"#FFF9F9":"transparent"}}>
                  <td style={tdS}><span style={{color:C.primary,fontWeight:600}}>{p.sku}</span></td>
                  <td style={tdS}><div style={{fontWeight:600}}>{p.name}</div></td>
                  <td style={tdS}><span style={{background:C.primaryL,color:C.primary,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:500}}>{p.category}</span></td>
                  <td style={tdS}>{p.warehouse}</td>
                  <td style={tdS}><span style={{fontWeight:700,fontSize:15,color:p.stock<=p.minStock?C.danger:C.success}}>{fmtN(p.stock)}</span></td>
                  <td style={tdS}><span style={{color:C.muted}}>{p.minStock}</span></td>
                  <td style={tdS}>{fmt(p.price)}</td>
                  <td style={tdS}><span style={{fontWeight:600}}>{fmt(p.stock*p.price)}</span></td>
                  <td style={tdS}><Badge status={p.stock<=p.minStock?"kritis":"aman"}/></td>
                  <td style={tdS}>
                    <div style={{display:"flex",gap:5}}>
                      <button onClick={()=>setEditProd({...p})} style={btnSm("#64748B")}>✏️ Edit</button>
                      <button onClick={()=>setConfirmDel(p.id)} style={btnSm(C.danger)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ── STOK MASUK ── */}
      {tab==="masuk"&&(
        <div style={{display:"grid",gridTemplateColumns:"360px 1fr",gap:16,alignItems:"start"}}>
          <div style={{...card,borderTop:`4px solid ${C.success}`}}>
            <div style={{fontWeight:700,fontSize:15,color:C.success,marginBottom:4}}>📥 Stok Masuk</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:18}}>Catat penerimaan barang dari supplier</div>
            <div style={{display:"grid",gap:12}}>
              <F label="Produk *"><select value={fMasuk.productId} onChange={e=>setFMasuk(p=>({...p,productId:e.target.value}))} style={sel}><option value="">— Pilih produk —</option>{products.map(p=><option key={p.id} value={p.id}>{p.name} (Stok: {p.stock})</option>)}</select></F>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <F label="Jumlah *"><input type="number" min="1" placeholder="0" value={fMasuk.qty} onChange={e=>setFMasuk(p=>({...p,qty:e.target.value}))} style={inp}/></F>
                <F label="Tanggal"><input type="date" value={fMasuk.date} onChange={e=>setFMasuk(p=>({...p,date:e.target.value}))} style={inp}/></F>
              </div>
              <F label="Supplier"><input type="text" placeholder="Nama supplier..." value={fMasuk.supplier} onChange={e=>setFMasuk(p=>({...p,supplier:e.target.value}))} style={inp}/></F>
              <F label="No. PO / Referensi"><input type="text" placeholder="PO-2024-xxx" value={fMasuk.ref} onChange={e=>setFMasuk(p=>({...p,ref:e.target.value}))} style={inp}/></F>
              <F label="Catatan"><input type="text" placeholder="Keterangan..." value={fMasuk.note} onChange={e=>setFMasuk(p=>({...p,note:e.target.value}))} style={inp}/></F>
              {fMasuk.productId&&fMasuk.qty&&<div style={{background:C.successL,borderRadius:8,padding:"9px 13px",fontSize:13,color:"#065F46"}}>Stok sesudah: <b>{(products.find(p=>p.id===Number(fMasuk.productId))?.stock||0)+Number(fMasuk.qty)} unit</b></div>}
              <button onClick={handleMasuk} style={{...btn(C.success),padding:"10px"}}>📥 Simpan Stok Masuk</button>
            </div>
          </div>
          <div style={card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontWeight:700,fontSize:15}}>Riwayat Stok Masuk</div>
              <ExportBtn onExcel={()=>{const e=buildMovExport(movements.filter(m=>m.type==="in"));exportExcel([{name:"Stok Masuk",...e}]);}} onPdf={()=>{const e=buildMovExport(movements.filter(m=>m.type==="in"));exportPDF("Stok Masuk",e.html);}}/>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.surface}}>{["Tgl","Produk","Qty","+/-","Ref","Catatan","Aksi"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
              <tbody>
                {movements.filter(m=>m.type==="in").length===0
                  ?<tr><td colSpan={7} style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada riwayat</td></tr>
                  :movements.filter(m=>m.type==="in").slice().reverse().map(m=>(
                    <tr key={m.id} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td style={tdS}>{m.date}</td>
                      <td style={tdS}><div style={{fontWeight:600}}>{m.product}</div></td>
                      <td style={tdS}><span style={{fontWeight:700,color:C.success}}>+{m.qty}</span></td>
                      <td style={tdS}><span style={{color:C.muted,fontSize:12}}>{m.before??"-"} → {m.after??"-"}</span></td>
                      <td style={tdS}><span style={{color:C.primary,fontWeight:600,fontSize:12}}>{m.ref}</span></td>
                      <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{m.note}</span></td>
                      <td style={tdS}>
                        <div style={{display:"flex",gap:5}}>
                          <button onClick={()=>setEditMov({...m})} style={btnSm("#64748B")}>✏️</button>
                          <button onClick={()=>setConfirmDelMov(m.id)} style={btnSm(C.danger)}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── STOK KELUAR ── */}
      {tab==="keluar"&&(
        <div style={{display:"grid",gridTemplateColumns:"360px 1fr",gap:16,alignItems:"start"}}>
          <div style={{...card,borderTop:`4px solid ${C.danger}`}}>
            <div style={{fontWeight:700,fontSize:15,color:C.danger,marginBottom:4}}>📤 Stok Keluar</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:18}}>Catat pengeluaran barang dari gudang</div>
            <div style={{display:"grid",gap:12}}>
              <F label="Produk *"><select value={fKeluar.productId} onChange={e=>setFKeluar(p=>({...p,productId:e.target.value}))} style={sel}><option value="">— Pilih produk —</option>{products.map(p=><option key={p.id} value={p.id}>{p.name} (Stok: {p.stock})</option>)}</select></F>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <F label="Jumlah *"><input type="number" min="1" placeholder="0" value={fKeluar.qty} onChange={e=>setFKeluar(p=>({...p,qty:e.target.value}))} style={inp}/></F>
                <F label="Tanggal"><input type="date" value={fKeluar.date} onChange={e=>setFKeluar(p=>({...p,date:e.target.value}))} style={inp}/></F>
              </div>
              <F label="Alasan"><select value={fKeluar.reason} onChange={e=>setFKeluar(p=>({...p,reason:e.target.value}))} style={sel}>{["penjualan","retur","rusak","transfer","lainnya"].map(r=><option key={r} value={r}>{r}</option>)}</select></F>
              <F label="Referensi"><input type="text" placeholder="SO-xxx" value={fKeluar.ref} onChange={e=>setFKeluar(p=>({...p,ref:e.target.value}))} style={inp}/></F>
              <F label="Catatan"><input type="text" placeholder="Keterangan..." value={fKeluar.note} onChange={e=>setFKeluar(p=>({...p,note:e.target.value}))} style={inp}/></F>
              {fKeluar.productId&&fKeluar.qty&&(()=>{const p=products.find(x=>x.id===Number(fKeluar.productId));const s=(p?.stock||0)-Number(fKeluar.qty);return<div style={{background:s<0?C.dangerL:C.surface,borderRadius:8,padding:"9px 13px",fontSize:13}}>{s<0?<span style={{color:C.danger,fontWeight:600}}>⚠️ Stok tidak cukup!</span>:<span>Stok sesudah: <b style={{color:s<=p?.minStock?C.danger:C.success}}>{s} unit</b></span>}</div>;})()}
              <button onClick={handleKeluar} style={{...btn(C.danger),padding:"10px"}}>📤 Simpan Stok Keluar</button>
            </div>
          </div>
          <div style={card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontWeight:700,fontSize:15}}>Riwayat Stok Keluar</div>
              <ExportBtn onExcel={()=>{const e=buildMovExport(movements.filter(m=>m.type==="out"));exportExcel([{name:"Stok Keluar",...e}]);}} onPdf={()=>{const e=buildMovExport(movements.filter(m=>m.type==="out"));exportPDF("Stok Keluar",e.html);}}/>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.surface}}>{["Tgl","Produk","Qty","+/-","Alasan","Ref","Catatan","Aksi"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
              <tbody>
                {movements.filter(m=>m.type==="out").length===0
                  ?<tr><td colSpan={8} style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada riwayat</td></tr>
                  :movements.filter(m=>m.type==="out").slice().reverse().map(m=>(
                    <tr key={m.id} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td style={tdS}>{m.date}</td>
                      <td style={tdS}><div style={{fontWeight:600}}>{m.product}</div></td>
                      <td style={tdS}><span style={{fontWeight:700,color:C.danger}}>-{m.qty}</span></td>
                      <td style={tdS}><span style={{color:C.muted,fontSize:12}}>{m.before??"-"} → {m.after??"-"}</span></td>
                      <td style={tdS}><span style={{background:C.surface,color:C.muted,padding:"2px 7px",borderRadius:10,fontSize:11}}>{m.subtype||"-"}</span></td>
                      <td style={tdS}><span style={{color:C.primary,fontWeight:600,fontSize:12}}>{m.ref}</span></td>
                      <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{m.note}</span></td>
                      <td style={tdS}>
                        <div style={{display:"flex",gap:5}}>
                          <button onClick={()=>setEditMov({...m})} style={btnSm("#64748B")}>✏️</button>
                          <button onClick={()=>setConfirmDelMov(m.id)} style={btnSm(C.danger)}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── STOK OPNAME ── */}
      {tab==="opname"&&(
        <div>
          {movements.filter(m=>m.subtype==="opname").length>0&&(
            <div style={{...card,marginBottom:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{fontWeight:700,fontSize:14}}>🔍 Riwayat Penyesuaian Opname</div>
                <ExportBtn
                  onExcel={()=>{
                    const opMoves=movements.filter(m=>m.subtype==="opname");
                    const cols=["Tanggal","Produk","Jenis","Selisih","Referensi","Catatan","Stok Sebelum","Stok Sesudah"];
                    const rows=opMoves.map(m=>[m.date,m.product,m.type==="in"?"Masuk":"Keluar",m.type==="in"?m.qty:-m.qty,m.ref,m.note,m.before??"-",m.after??"-"]);
                    exportExcel([{name:"Stok Opname",cols,rows}]);
                  }}
                  onPdf={()=>{
                    const opMoves=movements.filter(m=>m.subtype==="opname");
                    const totalPlus=opMoves.filter(m=>m.type==="in").reduce((s,m)=>s+m.qty,0);
                    const totalMinus=opMoves.filter(m=>m.type==="out").reduce((s,m)=>s+m.qty,0);
                    const html=`<h2>Riwayat Stok Opname</h2>
                      <table><thead><tr>${["Tanggal","Produk","Jenis","Selisih","Referensi","Catatan","Sebelum","Sesudah"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>
                      ${opMoves.map(m=>{const d=m.type==="in"?m.qty:-m.qty;return`<tr><td>${m.date}</td><td><b>${m.product}</b></td><td><span class="${m.type==="in"?"g":"r"}">${m.type==="in"?"Masuk":"Keluar"}</span></td><td><b style="color:${m.type==="in"?"#10B97B":"#EF4444"}">${d>0?"+":""}${d}</b></td><td style="color:#1B5FBF">${m.ref}</td><td style="color:#64748B">${m.note}</td><td>${m.before??"-"}</td><td><b>${m.after??"-"}</b></td></tr>`;}).join("")}
                      </tbody></table>
                      <div class="sum"><div><label>Total Penyesuaian</label><span>${opMoves.length}</span></div><div><label>Stok Bertambah</label><span style="color:#10B97B">+${totalPlus} unit</span></div><div><label>Stok Berkurang</label><span style="color:#EF4444">-${totalMinus} unit</span></div></div>`;
                    exportPDF("Laporan Stok Opname",html);
                  }}
                />
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:C.surface}}>{["Tanggal","Produk","Selisih","Referensi","Catatan","Aksi"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
                <tbody>
                  {movements.filter(m=>m.subtype==="opname").slice().reverse().map(m=>{
                    const diff=m.type==="in"?m.qty:-m.qty;
                    return(
                      <tr key={m.id} style={{borderBottom:`1px solid ${C.border}`}}>
                        <td style={tdS}>{m.date}</td>
                        <td style={tdS}><span style={{fontWeight:600}}>{m.product}</span></td>
                        <td style={tdS}><span style={{fontWeight:700,color:diff>0?C.success:C.danger}}>{diff>0?"+":""}{diff}</span></td>
                        <td style={tdS}><span style={{color:C.purple,fontWeight:600,fontSize:12}}>{m.ref}</span></td>
                        <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{m.note}</span></td>
                        <td style={tdS}><button onClick={()=>setConfirmDelMov(m.id)} style={btnSm(C.danger)}>🗑️</button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
          <div style={{...card,borderTop:`4px solid ${C.purple}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
              <div><div style={{fontWeight:700,fontSize:16,color:C.purple}}>🔍 Stok Opname</div><div style={{fontSize:12,color:C.muted,marginTop:2}}>Hitung fisik dan sesuaikan stok sistem</div></div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <div><label style={{...lbl,marginBottom:4}}>Tanggal</label><input type="date" value={opDate} onChange={e=>setOpDate(e.target.value)} style={{...inp,width:160}}/></div>
                {opRows.length===0?<button onClick={startOpname} style={{...btn(C.purple),marginTop:20}}>Mulai Opname</button>
                  :!opDone&&<button onClick={submitOpname} style={{...btn(C.success),marginTop:20}}>✓ Simpan</button>}
                {opRows.length>0&&<button onClick={()=>{setOpRows([]);setOpDone(false);}} style={{...btn("#94A3B8"),marginTop:20}}>Reset</button>}
              </div>
            </div>
            {opRows.length===0?<div style={{textAlign:"center",padding:"40px 0",color:C.muted}}><div style={{fontSize:40,marginBottom:12}}>🔍</div><div style={{fontSize:14,fontWeight:600}}>Belum ada sesi opname aktif</div></div>
            :(
              <>
                {opDone&&<div style={{background:C.successL,borderRadius:10,padding:"13px 17px",marginBottom:16,display:"flex",gap:10,alignItems:"center"}}><span style={{fontSize:24}}>✅</span><div><div style={{fontWeight:700,color:"#065F46"}}>Stok Opname Disimpan!</div><div style={{fontSize:12,color:"#065F46"}}>Penyesuaian tercatat ke mutasi stok.</div></div></div>}
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr style={{background:C.surface}}>{["Produk","Gudang","Stok Sistem","Stok Fisik","Selisih","Catatan"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
                  <tbody>
                    {opRows.map(r=>{
                      const diff=r.aktual===""?0:r.aktual-r.stock;
                      return(
                        <tr key={r.id} style={{borderBottom:`1px solid ${C.border}`,background:diff!==0?"#FAFFF6":"transparent"}}>
                          <td style={tdS}><div style={{fontWeight:600}}>{r.name}</div><div style={{fontSize:11,color:C.muted}}>{r.sku}</div></td>
                          <td style={tdS}>{r.warehouse}</td>
                          <td style={tdS}><span style={{fontWeight:700,fontSize:15}}>{r.stock}</span></td>
                          <td style={tdS}><input type="number" min="0" value={r.aktual} disabled={opDone} onChange={e=>setOpRows(prev=>prev.map(x=>x.id===r.id?{...x,aktual:e.target.value===""?"":Number(e.target.value)}:x))} style={{...inp,width:90,fontWeight:700,textAlign:"center",borderColor:diff>0?C.success:diff<0?C.danger:C.border,color:diff>0?C.success:diff<0?C.danger:C.text}}/></td>
                          <td style={tdS}>{r.aktual===""||diff===0?<span style={{color:C.muted}}>—</span>:<span style={{fontWeight:800,fontSize:15,color:diff>0?C.success:C.danger}}>{diff>0?"+":""}{diff}</span>}</td>
                          <td style={tdS}><input type="text" placeholder="Catatan..." value={r.note||""} disabled={opDone} onChange={e=>setOpRows(prev=>prev.map(x=>x.id===r.id?{...x,note:e.target.value}:x))} style={{...inp,width:180,fontSize:12}}/></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {!opDone&&<div style={{marginTop:14,display:"flex",justifyContent:"flex-end"}}><button onClick={submitOpname} style={{...btn(C.success),padding:"10px 24px"}}>✓ Simpan Semua Penyesuaian</button></div>}
                {opDone&&(
                  <div style={{marginTop:14,display:"flex",justifyContent:"flex-end",gap:8}}>
                    <ExportBtn
                      onExcel={()=>{
                        const cols=["Produk","SKU","Gudang","Stok Sistem","Stok Fisik","Selisih","Catatan"];
                        const rows=opRows.map(r=>[r.name,r.sku,r.warehouse,r.stock,r.aktual===""?r.stock:r.aktual,(r.aktual===""?0:r.aktual-r.stock),r.note||""]);
                        exportExcel([{name:"Stok Opname",cols,rows}]);
                      }}
                      onPdf={()=>{
                        const changed=opRows.filter(r=>r.aktual!==""&&r.aktual!==r.stock);
                        const html=`<h2>Berita Acara Stok Opname — ${opDate}</h2>
                          <table><thead><tr>${["Produk","SKU","Gudang","Stok Sistem","Stok Fisik","Selisih","Catatan"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>
                          ${opRows.map(r=>{const d=r.aktual===""?0:r.aktual-r.stock;return`<tr><td><b>${r.name}</b></td><td>${r.sku}</td><td>${r.warehouse}</td><td>${r.stock}</td><td><b style="color:${d>0?"#10B97B":d<0?"#EF4444":"inherit"}">${r.aktual===""?r.stock:r.aktual}</b></td><td><b style="color:${d>0?"#10B97B":d<0?"#EF4444":"#64748B"}">${d===0?"—":(d>0?"+":"")+d}</b></td><td style="color:#64748B">${r.note||""}</td></tr>`;}).join("")}
                          </tbody></table>
                          <div class="sum"><div><label>Total Produk</label><span>${opRows.length}</span></div><div><label>Ada Selisih</label><span style="color:#F59E0B">${changed.length}</span></div><div><label>Bertambah</label><span style="color:#10B97B">+${changed.filter(r=>r.aktual>r.stock).reduce((s,r)=>s+(r.aktual-r.stock),0)} unit</span></div><div><label>Berkurang</label><span style="color:#EF4444">-${changed.filter(r=>r.aktual<r.stock).reduce((s,r)=>s+(r.stock-r.aktual),0)} unit</span></div></div>`;
                        exportPDF(`Berita Acara Stok Opname ${opDate}`,html);
                      }}
                    />
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      )}

      {/* Alerts */}
      {alert&&<Modal onClose={()=>setAlert(null)}><div style={{textAlign:"center",padding:"20px 0"}}><div style={{fontSize:48,marginBottom:12}}>{alert.ok?"✅":"❌"}</div><div style={{fontSize:15,fontWeight:600,color:alert.ok?C.success:C.danger}}>{alert.msg}</div><button onClick={()=>setAlert(null)} style={{...btn(C.primary),marginTop:20}}>OK</button></div></Modal>}

      {/* Edit Product Modal */}
      {editProd&&<Modal onClose={()=>setEditProd(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editProd.id?"✏️ Edit Produk":"➕ Tambah Produk"}</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="SKU *"><input value={editProd.sku} onChange={e=>setEditProd(p=>({...p,sku:e.target.value}))} style={inp} placeholder="PRD-xxx"/></F>
            <F label="Nama Produk *"><input value={editProd.name} onChange={e=>setEditProd(p=>({...p,name:e.target.value}))} style={inp} placeholder="Nama..."/></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Kategori"><input value={editProd.category} onChange={e=>setEditProd(p=>({...p,category:e.target.value}))} style={inp} placeholder="Elektronik"/></F>
            <F label="Gudang"><select value={editProd.warehouse} onChange={e=>setEditProd(p=>({...p,warehouse:e.target.value}))} style={sel}>{["Gudang A","Gudang B","Gudang C"].map(g=><option key={g}>{g}</option>)}</select></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
            <F label="Stok"><input type="number" min="0" value={editProd.stock} onChange={e=>setEditProd(p=>({...p,stock:Number(e.target.value)}))} style={inp}/></F>
            <F label="Min Stok"><input type="number" min="0" value={editProd.minStock} onChange={e=>setEditProd(p=>({...p,minStock:Number(e.target.value)}))} style={inp}/></F>
            <F label="Harga (IDR)"><input type="number" min="0" value={editProd.price} onChange={e=>setEditProd(p=>({...p,price:Number(e.target.value)}))} style={inp}/></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Unit"><input value={editProd.unit} onChange={e=>setEditProd(p=>({...p,unit:e.target.value}))} style={inp} placeholder="Unit"/></F>
          </div>
          <button onClick={()=>saveProduct(editProd.id?editProd:{...editProd,id:uid()})} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* Edit Movement Modal */}
      {editMov&&<Modal onClose={()=>setEditMov(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>✏️ Edit Mutasi Stok</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Tanggal"><input type="date" value={editMov.date} onChange={e=>setEditMov(m=>({...m,date:e.target.value}))} style={inp}/></F>
            <F label="Referensi"><input value={editMov.ref} onChange={e=>setEditMov(m=>({...m,ref:e.target.value}))} style={inp}/></F>
          </div>
          <F label="Catatan"><input value={editMov.note} onChange={e=>setEditMov(m=>({...m,note:e.target.value}))} style={inp}/></F>
          <button onClick={()=>saveMovement(editMov)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {confirmDel&&<ConfirmModal msg={`Hapus produk "${products.find(p=>p.id===confirmDel)?.name}"? Semua mutasi terkait juga dihapus.`} onConfirm={()=>deleteProduct(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
      {confirmDelMov&&<ConfirmModal msg="Hapus catatan mutasi ini?" onConfirm={()=>deleteMovement(confirmDelMov)} onCancel={()=>setConfirmDelMov(null)}/>}
    </div>
  );
}

// ─── CASH FLOW PAGE ───────────────────────────────────────────────────────────
function CashFlowPage({invoices,manualCF,setManualCF}){
  const [filterMonth,setFilterMonth]=useState("2024-06");
  const [editCF,setEditCF]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [showAdd,setShowAdd]=useState(false);
  const [newE,setNewE]=useState({date:"",type:"out",category:"Operasional",description:"",amount:"",ref:""});

  const allEntries=useMemo(()=>{
    const fromInv=invoices.filter(i=>i.status==="paid"&&i.paidDate).map(i=>({id:`CF-INV-${i.id}`,date:i.paidDate,type:"in",category:"Penjualan",description:`Pembayaran invoice ${i.id} – ${i.customer}`,amount:i.grandTotal,ref:i.id,source:"invoice"}));
    return[...fromInv,...manualCF.map(m=>({...m,source:"manual"}))].sort((a,b)=>a.date.localeCompare(b.date));
  },[invoices,manualCF]);

  const filtered=useMemo(()=>filterMonth?allEntries.filter(e=>e.date.startsWith(filterMonth)):allEntries,[allEntries,filterMonth]);
  const totalIn=filtered.filter(e=>e.type==="in").reduce((s,e)=>s+e.amount,0);
  const totalOut=filtered.filter(e=>e.type==="out").reduce((s,e)=>s+e.amount,0);
  const net=totalIn-totalOut;
  let run=0;
  const withBal=filtered.map(e=>{run+=e.type==="in"?e.amount:-e.amount;return{...e,balance:run};});
  const catBreak=filtered.filter(e=>e.type==="out").reduce((a,e)=>{a[e.category]=(a[e.category]||0)+e.amount;return a;},{});

  const handleAdd=()=>{
    if(!newE.date||!newE.description||!newE.amount)return;
    setManualCF(prev=>[...prev,{id:`CF-M-${uid()}`,date:newE.date,type:newE.type,category:newE.category,description:newE.description,amount:Number(newE.amount),ref:newE.ref||"-"}]);
    setNewE({date:"",type:"out",category:"Operasional",description:"",amount:"",ref:""});
    setShowAdd(false);
  };
  const saveCF=(e)=>{setManualCF(prev=>prev.map(x=>x.id===e.id?e:x));setEditCF(null);};
  const deleteCF=(id)=>{setManualCF(prev=>prev.filter(x=>x.id!==id));setConfirmDel(null);};

  const months=["2024-06","2024-05","2024-04","Semua"];

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total Pemasukan",v:fmt(totalIn),c:C.success,i:"📥",sub:`${filtered.filter(e=>e.type==="in").length} transaksi`},
          {l:"Total Pengeluaran",v:fmt(totalOut),c:C.danger,i:"📤",sub:`${filtered.filter(e=>e.type==="out").length} transaksi`},
          {l:"Arus Kas Bersih",v:fmt(net),c:net>=0?C.success:C.danger,i:net>=0?"📈":"📉",sub:net>=0?"Surplus":"Defisit"},
          {l:"Dari Invoice",v:fmt(filtered.filter(e=>e.source==="invoice").reduce((s,e)=>s+e.amount,0)),c:C.primary,i:"🧾",sub:"Auto tersinkron"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{k.sub}</div></div>
              <span style={{fontSize:22}}>{k.i}</span>
            </div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 280px",gap:16,marginBottom:16}}>
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>Tren Arus Kas</div>
          <div style={{display:"flex",gap:6,marginBottom:18}}>{months.map(m=><button key={m} onClick={()=>setFilterMonth(m==="Semua"?"":m)} style={{padding:"5px 12px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:(filterMonth===m||(m==="Semua"&&!filterMonth))?C.primary:C.surface,color:(filterMonth===m||(m==="Semua"&&!filterMonth))?"#FFF":C.muted}}>{m}</button>)}</div>
          {[["Pemasukan",totalIn,C.success],["Pengeluaran",totalOut,C.danger]].map(([label,val,color])=>{
            const pct=Math.round((val/Math.max(totalIn,totalOut,1))*100);
            return(<div key={label} style={{marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5,fontSize:12}}><span style={{color:C.muted,fontWeight:500}}>{label}</span><span style={{fontWeight:700,color}}>{fmt(val)}</span></div><div style={{background:C.surface,borderRadius:6,height:10,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:6}}/></div></div>);
          })}
          <div style={{background:net>=0?C.successL:C.dangerL,borderRadius:10,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:600,fontSize:13,color:net>=0?"#065F46":"#991B1B"}}>Arus Kas Bersih</span>
            <span style={{fontWeight:800,fontSize:18,color:net>=0?C.success:C.danger}}>{net>=0?"+":""}{fmt(net)}</span>
          </div>
        </div>
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>Pengeluaran per Kategori</div>
          {Object.entries(catBreak).sort((a,b)=>b[1]-a[1]).map(([cat,amt])=>{const pct=Math.round((amt/totalOut)*100);return(<div key={cat} style={{marginBottom:13}}><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}><span style={{fontWeight:600}}>{cat}</span><span style={{color:C.muted}}>{pct}%</span></div><div style={{background:C.surface,borderRadius:4,height:7}}><div style={{width:`${pct}%`,height:"100%",background:C.danger,borderRadius:4}}/></div></div>);})}
        </div>
      </div>
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontWeight:700,fontSize:15}}>Riwayat Transaksi Kas</div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setShowAdd(true)} style={btn(C.purple)}>+ Tambah</button>
            <ExportBtn onExcel={()=>{const e=buildCFExport(withBal);exportExcel([{name:"Arus Kas",...e}]);}} onPdf={()=>{const e=buildCFExport(withBal);exportPDF("Arus Kas",e.html);}}/>
          </div>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:C.surface}}>{["Tgl","Keterangan","Kategori","Ref","Sumber","Jenis","Jumlah","Saldo","Aksi"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>
            {withBal.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:32,color:C.muted}}>Tidak ada transaksi</td></tr>
              :withBal.map(e=>(
                <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`,background:e.source==="invoice"?"#FAFFFE":"transparent"}}>
                  <td style={tdS}>{e.date}</td>
                  <td style={tdS}><div style={{fontWeight:600}}>{e.description}</div>{e.source==="invoice"&&<div style={{fontSize:11,color:C.primary,fontWeight:500}}>🔗 Auto Invoice</div>}</td>
                  <td style={tdS}><span style={{background:e.category==="Penjualan"?C.primaryL:C.surface,color:e.category==="Penjualan"?C.primary:C.muted,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:500}}>{e.category}</span></td>
                  <td style={tdS}><span style={{color:C.primary,fontWeight:600,fontSize:12}}>{e.ref}</span></td>
                  <td style={tdS}><span style={{fontSize:11,color:e.source==="invoice"?C.success:e.source==="payroll"?C.warning:C.muted}}>{e.source==="invoice"?"🧾 Invoice":e.source==="payroll"?"💸 Payroll":"✏️ Manual"}</span></td>
                  <td style={tdS}><Badge status={e.type}/></td>
                  <td style={tdS}><span style={{fontWeight:700,color:e.type==="in"?C.success:C.danger}}>{e.type==="in"?"+":"-"}{fmt(e.amount)}</span></td>
                  <td style={tdS}><span style={{fontWeight:700,color:e.balance>=0?C.text:C.danger}}>{fmt(e.balance)}</span></td>
                  <td style={tdS}>
                    {e.source==="manual"?(
                      <div style={{display:"flex",gap:5}}>
                        <button onClick={()=>setEditCF({...e})} style={btnSm("#64748B")}>✏️</button>
                        <button onClick={()=>setConfirmDel(e.id)} style={btnSm(C.danger)}>🗑️</button>
                      </div>
                    ):<span style={{fontSize:11,color:C.muted}}>—</span>}                  </td>
                </tr>
              ))
            }
          </tbody>
        </table>
        <div style={{marginTop:14,padding:"13px 16px",background:C.surface,borderRadius:10,display:"flex",gap:28,borderTop:`2px solid ${C.border}`}}>
          <div><div style={{fontSize:11,color:C.muted}}>Total Pemasukan</div><div style={{fontWeight:700,color:C.success}}>{fmt(totalIn)}</div></div>
          <div><div style={{fontSize:11,color:C.muted}}>Total Pengeluaran</div><div style={{fontWeight:700,color:C.danger}}>{fmt(totalOut)}</div></div>
          <div style={{marginLeft:"auto"}}><div style={{fontSize:11,color:C.muted}}>Saldo Akhir</div><div style={{fontWeight:800,fontSize:16,color:net>=0?C.success:C.danger}}>{fmt(net)}</div></div>
        </div>
      </div>

      {/* Add / Edit CF Modal */}
      {(showAdd||editCF)&&<Modal onClose={()=>{setShowAdd(false);setEditCF(null);}}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editCF?"✏️ Edit Transaksi":"➕ Tambah Transaksi"}</div>
        <div style={{display:"grid",gap:13}}>
          {[{l:"Tanggal",k:"date",t:"date"},{l:"Keterangan",k:"description",t:"text",ph:"Deskripsi..."},{l:"Jumlah (IDR)",k:"amount",t:"number",ph:"0"},{l:"Referensi",k:"ref",t:"text",ph:"No. dokumen"}].map(f=>(
            <F key={f.k} label={f.l}><input type={f.t} placeholder={f.ph} value={editCF?editCF[f.k]:newE[f.k]} onChange={e=>editCF?setEditCF(p=>({...p,[f.k]:e.target.value})):setNewE(p=>({...p,[f.k]:e.target.value}))} style={inp}/></F>
          ))}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Jenis"><select value={editCF?editCF.type:newE.type} onChange={e=>editCF?setEditCF(p=>({...p,type:e.target.value})):setNewE(p=>({...p,type:e.target.value}))} style={sel}><option value="in">Pemasukan</option><option value="out">Pengeluaran</option></select></F>
            <F label="Kategori"><select value={editCF?editCF.category:newE.category} onChange={e=>editCF?setEditCF(p=>({...p,category:e.target.value})):setNewE(p=>({...p,category:e.target.value}))} style={sel}>{["Operasional","Logistik","Utilitas","Gaji","Pembelian","Lain-lain","Penjualan"].map(c=><option key={c}>{c}</option>)}</select></F>
          </div>
          <button onClick={()=>editCF?saveCF(editCF):handleAdd()} style={{...btn(C.purple),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg="Hapus transaksi kas ini?" onConfirm={()=>deleteCF(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

const buildEmpExport = (emps)=>({
  cols:["NIK","Nama","Departemen","Jabatan","Status","Tgl Masuk","Gaji Pokok","T.Makan","T.Transport","T.Jabatan","Pot.BPJS","Pot.PPh","Bank","No.Rek"],
  rows:emps.map(e=>[e.nik,e.name,e.dept,e.position,e.status,e.joinDate,fmtR(e.gajiPokok),fmtR(e.tunjMakan),fmtR(e.tunjTransport),fmtR(e.tunjJabatan),fmtR(e.potonganBPJS),fmtR(e.potonganPPh),e.bank,e.noRek]),
  html:`<h2>Data Karyawan</h2><table><thead><tr>${["NIK","Nama","Dept","Jabatan","Status","Gaji Pokok","Take Home"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${emps.map(e=>`<tr><td>${e.nik}</td><td><b>${e.name}</b></td><td>${e.dept}</td><td>${e.position}</td><td><span class="${e.status==="aktif"?"g":"r"}">${e.status}</span></td><td>${fmt(e.gajiPokok)}</td><td><b>${fmt(e.gajiPokok+e.tunjMakan+e.tunjTransport+e.tunjJabatan-e.potonganBPJS-e.potonganPPh)}</b></td></tr>`).join("")}</tbody></table><div class="sum"><div><label>Total Karyawan</label><span>${emps.length}</span></div><div><label>Aktif</label><span style="color:#10B97B">${emps.filter(e=>e.status==="aktif").length}</span></div><div><label>Total Gaji Pokok</label><span>${fmt(emps.reduce((s,e)=>s+e.gajiPokok,0))}</span></div></div>`
});
const buildPayrollExport = (pays)=>({
  cols:["ID","Nama","Bulan","Gaji Pokok","T.Makan","T.Transport","T.Jabatan","Lembur","Total Tunjangan","Pot.BPJS","Pot.PPh","Pot.Lain","Total Potongan","Take Home","Status","Tgl Bayar"],
  rows:pays.map(p=>[p.id,p.employeeName,p.month,fmtR(p.gajiPokok),fmtR(p.tunjMakan),fmtR(p.tunjTransport),fmtR(p.tunjJabatan),fmtR(p.lemburAmount),fmtR(p.totalTunjangan),fmtR(p.potonganBPJS),fmtR(p.potonganPPh),fmtR(p.potonganLain),fmtR(p.totalPotongan),fmtR(p.takehome),p.status,p.payDate||"-"]),
  html:`<h2>Rekap Penggajian</h2><table><thead><tr>${["Nama","Bulan","Gaji Pokok","Total Tunjangan","Total Potongan","Take Home","Status"].map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${pays.map(p=>`<tr><td><b>${p.employeeName}</b></td><td>${p.month}</td><td>${fmt(p.gajiPokok)}</td><td style="color:#10B97B">+${fmt(p.totalTunjangan)}</td><td style="color:#EF4444">-${fmt(p.totalPotongan)}</td><td><b>${fmt(p.takehome)}</b></td><td><span class="${p.status==="dibayar"?"g":"r"}">${p.status}</span></td></tr>`).join("")}</tbody></table><div class="sum"><div><label>Total Dibayar</label><span>${pays.length}</span></div><div><label>Total Take Home</label><span>${fmt(pays.reduce((s,p)=>s+p.takehome,0))}</span></div><div><label>Sudah Dibayar</label><span style="color:#10B97B">${pays.filter(p=>p.status==="dibayar").length}</span></div></div>`
});

// ─── HR PAGE ──────────────────────────────────────────────────────────────────
function HRPage({employees,setEmployees,payrolls,setPayrolls}){
  const [tab,setTab]=useState("karyawan");
  const [editEmp,setEditEmp]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [editPay,setEditPay]=useState(null);
  const [confirmDelPay,setConfirmDelPay]=useState(null);
  const [slipModal,setSlipModal]=useState(null);
  const [filterMonth,setFilterMonth]=useState("");
  // New payroll form
  const [newPay,setNewPay]=useState({employeeId:"",month:today().slice(0,7),lemburJam:0,potonganLain:0,potonganLainKet:""});

  const saveEmployee=(e)=>{
    if(e._new){setEmployees(prev=>[...prev,{...e,id:uid(),_new:undefined}]);}
    else{setEmployees(prev=>prev.map(x=>x.id===e.id?e:x));}
    setEditEmp(null);
  };
  const deleteEmployee=(id)=>{setEmployees(prev=>prev.filter(e=>e.id!==id));setPayrolls(prev=>prev.filter(p=>p.employeeId!==id));setConfirmDel(null);};

  const hitungGaji=(emp,lemburJam,potonganLain)=>{
    const tarifLembur=Math.round(emp.gajiPokok/173);
    const lemburAmount=lemburJam*tarifLembur*1.5;
    const totalTunjangan=emp.tunjMakan+emp.tunjTransport+emp.tunjJabatan+lemburAmount;
    const totalPotongan=emp.potonganBPJS+emp.potonganPPh+potonganLain;
    const takehome=emp.gajiPokok+totalTunjangan-totalPotongan;
    return{lemburAmount,totalTunjangan,totalPotongan,takehome};
  };

  const prosesGaji=()=>{
    const emp=employees.find(e=>e.id===Number(newPay.employeeId));
    if(!emp||!newPay.month)return;
    const exists=payrolls.find(p=>p.employeeId===emp.id&&p.month===newPay.month);
    if(exists){alert("Gaji bulan ini sudah diproses!");return;}
    const lemburJam=Number(newPay.lemburJam)||0;
    const potonganLain=Number(newPay.potonganLain)||0;
    const {lemburAmount,totalTunjangan,totalPotongan,takehome}=hitungGaji(emp,lemburJam,potonganLain);
    const pay={
      id:`PAY-${newPay.month}-${String(uid()).slice(-4)}`,employeeId:emp.id,employeeName:emp.name,
      month:newPay.month,gajiPokok:emp.gajiPokok,tunjMakan:emp.tunjMakan,tunjTransport:emp.tunjTransport,
      tunjJabatan:emp.tunjJabatan,lemburJam,lemburAmount,potonganBPJS:emp.potonganBPJS,potonganPPh:emp.potonganPPh,
      potonganLain,potonganLainKet:newPay.potonganLainKet||"",totalTunjangan,totalPotongan,takehome,
      status:"proses",payDate:null,
    };
    setPayrolls(prev=>[...prev,pay]);
    setNewPay({employeeId:"",month:today().slice(0,7),lemburJam:0,potonganLain:0,potonganLainKet:""});
    setTab("penggajian");
  };

  const bayarGaji=(id)=>setPayrolls(prev=>prev.map(p=>p.id===id?{...p,status:"dibayar",payDate:today()}:p));
  const savePayroll=(p)=>{setPayrolls(prev=>prev.map(x=>x.id===p.id?p:x));setEditPay(null);};
  const deletePayroll=(id)=>{setPayrolls(prev=>prev.filter(p=>p.id!==id));setConfirmDelPay(null);};

  const filteredPays=filterMonth?payrolls.filter(p=>p.month===filterMonth):payrolls;
  const totalTH=filteredPays.reduce((s,p)=>s+p.takehome,0);
  const months=[...new Set(payrolls.map(p=>p.month))].sort().reverse();

  const tabs=[{id:"karyawan",l:"👤 Data Karyawan"},{id:"proses",l:"⚙️ Proses Gaji"},{id:"penggajian",l:"💸 Rekap Gaji"}];

  return(
    <div>
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total Karyawan",v:employees.length,c:C.primary,i:"👥"},
          {l:"Karyawan Aktif",v:employees.filter(e=>e.status==="aktif").length,c:C.success,i:"✅"},
          {l:"Total Gaji Bulan Ini",v:fmt(payrolls.filter(p=>p.month===today().slice(0,7)).reduce((s,p)=>s+p.takehome,0)),c:C.warning,i:"💰"},
          {l:"Belum Dibayar",v:payrolls.filter(p=>p.status!=="dibayar").length,c:C.danger,i:"⏳"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div></div>
              <span style={{fontSize:24}}>{k.i}</span>
            </div>
          </div>
        ))}
      </div>

      {/* Tabs */}
      <div style={{display:"flex",gap:4,marginBottom:16,background:"#FFF",borderRadius:12,padding:6,border:`1px solid ${C.border}`,width:"fit-content",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        {tabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?C.primary:"transparent",color:tab===t.id?"#FFF":C.muted}}>{t.l}</button>)}
      </div>

      {/* ── DATA KARYAWAN ── */}
      {tab==="karyawan"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div style={{fontWeight:700,fontSize:15}}>Data Karyawan</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setEditEmp({_new:true,nik:"",name:"",dept:"",position:"",status:"aktif",joinDate:today(),gajiPokok:0,tunjMakan:0,tunjTransport:0,tunjJabatan:0,potonganBPJS:0,potonganPPh:0,bank:"",noRek:""})} style={btn(C.primary)}>+ Tambah Karyawan</button>
              <ExportBtn onExcel={()=>{const e=buildEmpExport(employees);exportExcel([{name:"Karyawan",...e}]);}} onPdf={()=>{const e=buildEmpExport(employees);exportPDF("Data Karyawan",e.html);}}/>
            </div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["NIK","Nama","Departemen","Jabatan","Status","Gaji Pokok","Take Home Est.","Aksi"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
            <tbody>
              {employees.map(e=>{
                const th=e.gajiPokok+e.tunjMakan+e.tunjTransport+e.tunjJabatan-e.potonganBPJS-e.potonganPPh;
                return(
                  <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`}}>
                    <td style={tdS}><span style={{color:C.primary,fontWeight:600}}>{e.nik}</span></td>
                    <td style={tdS}><div style={{fontWeight:700}}>{e.name}</div><div style={{fontSize:11,color:C.muted}}>{e.joinDate}</div></td>
                    <td style={tdS}><span style={{background:C.primaryL,color:C.primary,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:500}}>{e.dept}</span></td>
                    <td style={tdS}>{e.position}</td>
                    <td style={tdS}><Badge status={e.status}/></td>
                    <td style={tdS}><span style={{fontWeight:600}}>{fmt(e.gajiPokok)}</span></td>
                    <td style={tdS}><span style={{fontWeight:700,color:C.success}}>{fmt(th)}</span></td>
                    <td style={tdS}>
                      <div style={{display:"flex",gap:5}}>
                        <button onClick={()=>setEditEmp({...e})} style={btnSm("#64748B")}>✏️ Edit</button>
                        <button onClick={()=>setConfirmDel(e.id)} style={btnSm(C.danger)}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* ── PROSES GAJI ── */}
      {tab==="proses"&&(
        <div style={{display:"grid",gridTemplateColumns:"400px 1fr",gap:16,alignItems:"start"}}>
          <div style={{...card,borderTop:`4px solid ${C.warning}`}}>
            <div style={{fontWeight:700,fontSize:15,color:C.warning,marginBottom:4}}>⚙️ Proses Penggajian</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:18}}>Hitung gaji + tunjangan + potongan</div>
            <div style={{display:"grid",gap:12}}>
              <F label="Karyawan *">
                <select value={newPay.employeeId} onChange={e=>setNewPay(p=>({...p,employeeId:e.target.value}))} style={sel}>
                  <option value="">— Pilih karyawan —</option>
                  {employees.filter(e=>e.status==="aktif").map(e=><option key={e.id} value={e.id}>{e.name} – {e.position}</option>)}
                </select>
              </F>
              <F label="Bulan Penggajian"><input type="month" value={newPay.month} onChange={e=>setNewPay(p=>({...p,month:e.target.value}))} style={inp}/></F>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <F label="Jam Lembur"><input type="number" min="0" value={newPay.lemburJam} onChange={e=>setNewPay(p=>({...p,lemburJam:e.target.value}))} style={inp} placeholder="0"/></F>
                <F label="Potongan Lain (IDR)"><input type="number" min="0" value={newPay.potonganLain} onChange={e=>setNewPay(p=>({...p,potonganLain:e.target.value}))} style={inp} placeholder="0"/></F>
              </div>
              <F label="Keterangan Potongan"><input type="text" value={newPay.potonganLainKet} onChange={e=>setNewPay(p=>({...p,potonganLainKet:e.target.value}))} style={inp} placeholder="Mis: kasbon, dll"/></F>

              {/* Preview kalkulasi */}
              {newPay.employeeId&&(()=>{
                const emp=employees.find(e=>e.id===Number(newPay.employeeId));
                if(!emp)return null;
                const {lemburAmount,totalTunjangan,totalPotongan,takehome}=hitungGaji(emp,Number(newPay.lemburJam)||0,Number(newPay.potonganLain)||0);
                return(
                  <div style={{background:C.surface,borderRadius:10,padding:14,fontSize:13}}>
                    <div style={{fontWeight:700,marginBottom:10,color:C.text}}>Preview Slip Gaji</div>
                    {[
                      ["Gaji Pokok",fmt(emp.gajiPokok),C.text],
                      ["Tunjangan Makan",fmt(emp.tunjMakan),C.success],
                      ["Tunjangan Transport",fmt(emp.tunjTransport),C.success],
                      ["Tunjangan Jabatan",fmt(emp.tunjJabatan),C.success],
                      ...(lemburAmount>0?[["Lembur ("+newPay.lemburJam+" jam)",fmt(lemburAmount),C.success]]:[]),
                      ["",null],
                      ["Pot. BPJS",fmt(emp.potonganBPJS),C.danger],
                      ["Pot. PPh21",fmt(emp.potonganPPh),C.danger],
                      ...(Number(newPay.potonganLain)>0?[["Pot. Lain",fmt(Number(newPay.potonganLain)),C.danger]]:[]),
                    ].map(([k,v,color],i)=>k===""?<div key={i} style={{height:1,background:C.border,margin:"6px 0"}}/>:(
                      <div key={i} style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                        <span style={{color:C.muted}}>{k}</span><span style={{color,fontWeight:500}}>{v}</span>
                      </div>
                    ))}
                    <div style={{borderTop:`2px solid ${C.border}`,marginTop:8,paddingTop:8,display:"flex",justifyContent:"space-between"}}>
                      <span style={{fontWeight:700}}>Take Home Pay</span>
                      <span style={{fontWeight:800,fontSize:16,color:C.primary}}>{fmt(takehome)}</span>
                    </div>
                  </div>
                );
              })()}
              <button onClick={prosesGaji} style={{...btn(C.warning),padding:"11px",color:"#FFF"}}>⚙️ Proses Penggajian</button>
            </div>
          </div>

          {/* Rekap cepat per karyawan */}
          <div style={card}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>Komponen Gaji Karyawan</div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.surface}}>{["Nama","Gaji Pokok","Tunjangan","BPJS","PPh21","Take Home"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
              <tbody>
                {employees.filter(e=>e.status==="aktif").map(e=>{
                  const tj=e.tunjMakan+e.tunjTransport+e.tunjJabatan;
                  const th=e.gajiPokok+tj-e.potonganBPJS-e.potonganPPh;
                  return(
                    <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td style={tdS}><div style={{fontWeight:600}}>{e.name}</div><div style={{fontSize:11,color:C.muted}}>{e.position}</div></td>
                      <td style={tdS}>{fmt(e.gajiPokok)}</td>
                      <td style={tdS}><span style={{color:C.success,fontWeight:500}}>+{fmt(tj)}</span></td>
                      <td style={tdS}><span style={{color:C.danger,fontSize:12}}>-{fmt(e.potonganBPJS)}</span></td>
                      <td style={tdS}><span style={{color:C.danger,fontSize:12}}>-{fmt(e.potonganPPh)}</span></td>
                      <td style={tdS}><span style={{fontWeight:700,color:C.primary}}>{fmt(th)}</span></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── REKAP PENGGAJIAN ── */}
      {tab==="penggajian"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <div style={{fontWeight:700,fontSize:15}}>Rekap Penggajian</div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <select value={filterMonth} onChange={e=>setFilterMonth(e.target.value)} style={{...sel,width:160}}>
                <option value="">Semua Bulan</option>
                {months.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
              <ExportBtn onExcel={()=>{const e=buildPayrollExport(filteredPays);exportExcel([{name:"Penggajian",...e}]);}} onPdf={()=>{const e=buildPayrollExport(filteredPays);exportPDF(`Rekap Penggajian ${filterMonth||"Semua"}`,e.html);}}/>
            </div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["ID","Nama","Bulan","Gaji Pokok","Tunjangan","Potongan","Take Home","Status","Aksi"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`}}>{h}</th>)}</tr></thead>
            <tbody>
              {filteredPays.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada data penggajian</td></tr>
                :filteredPays.slice().reverse().map(p=>(
                <tr key={p.id} style={{borderBottom:`1px solid ${C.border}`}}>
                  <td style={tdS}><span style={{color:C.primary,fontWeight:600,fontSize:12}}>{p.id}</span></td>
                  <td style={tdS}><div style={{fontWeight:600}}>{p.employeeName}</div></td>
                  <td style={tdS}>{p.month}</td>
                  <td style={tdS}>{fmt(p.gajiPokok)}</td>
                  <td style={tdS}><span style={{color:C.success,fontWeight:500}}>+{fmt(p.totalTunjangan)}</span></td>
                  <td style={tdS}><span style={{color:C.danger,fontWeight:500}}>-{fmt(p.totalPotongan)}</span></td>
                  <td style={tdS}><span style={{fontWeight:700,color:C.primary}}>{fmt(p.takehome)}</span></td>
                  <td style={tdS}><Badge status={p.status}/></td>
                  <td style={tdS}>
                    <div style={{display:"flex",gap:5}}>
                      <button onClick={()=>setSlipModal(p)} style={btnSm(C.primary)}>📄 Slip</button>
                      {p.status!=="dibayar"&&<button onClick={()=>bayarGaji(p.id)} style={btnSm(C.success)}>✓ Bayar</button>}
                      <button onClick={()=>setEditPay({...p})} style={btnSm("#64748B")}>✏️</button>
                      <button onClick={()=>setConfirmDelPay(p.id)} style={btnSm(C.danger)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{marginTop:14,padding:"13px 16px",background:C.surface,borderRadius:10,display:"flex",gap:28,borderTop:`2px solid ${C.border}`}}>
            <div><div style={{fontSize:11,color:C.muted}}>Total Karyawan</div><div style={{fontWeight:700}}>{filteredPays.length}</div></div>
            <div><div style={{fontSize:11,color:C.muted}}>Total Take Home</div><div style={{fontWeight:700,color:C.primary}}>{fmt(totalTH)}</div></div>
            <div><div style={{fontSize:11,color:C.muted}}>Sudah Dibayar</div><div style={{fontWeight:700,color:C.success}}>{fmt(filteredPays.filter(p=>p.status==="dibayar").reduce((s,p)=>s+p.takehome,0))}</div></div>
            <div><div style={{fontSize:11,color:C.muted}}>Belum Dibayar</div><div style={{fontWeight:700,color:C.danger}}>{fmt(filteredPays.filter(p=>p.status!=="dibayar").reduce((s,p)=>s+p.takehome,0))}</div></div>
          </div>
        </div>
      )}

      {/* Edit Employee Modal */}
      {editEmp&&<Modal onClose={()=>setEditEmp(null)} wide>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editEmp._new?"➕ Tambah Karyawan":"✏️ Edit Karyawan"}</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
            <F label="NIK"><input value={editEmp.nik} onChange={e=>setEditEmp(p=>({...p,nik:e.target.value}))} style={inp} placeholder="EMP-xxx"/></F>
            <F label="Nama Lengkap *"><input value={editEmp.name} onChange={e=>setEditEmp(p=>({...p,name:e.target.value}))} style={inp}/></F>
            <F label="Status"><select value={editEmp.status} onChange={e=>setEditEmp(p=>({...p,status:e.target.value}))} style={sel}><option value="aktif">Aktif</option><option value="nonaktif">Non-Aktif</option></select></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
            <F label="Departemen"><input value={editEmp.dept} onChange={e=>setEditEmp(p=>({...p,dept:e.target.value}))} style={inp} placeholder="IT, Keuangan..."/></F>
            <F label="Jabatan"><input value={editEmp.position} onChange={e=>setEditEmp(p=>({...p,position:e.target.value}))} style={inp}/></F>
            <F label="Tanggal Masuk"><input type="date" value={editEmp.joinDate} onChange={e=>setEditEmp(p=>({...p,joinDate:e.target.value}))} style={inp}/></F>
          </div>
          <div style={{background:C.surface,borderRadius:8,padding:14}}>
            <div style={{fontWeight:600,fontSize:13,marginBottom:10,color:C.text}}>💰 Komponen Gaji</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
              <F label="Gaji Pokok (IDR)"><input type="number" min="0" value={editEmp.gajiPokok} onChange={e=>setEditEmp(p=>({...p,gajiPokok:Number(e.target.value)}))} style={inp}/></F>
              <F label="T. Makan (IDR)"><input type="number" min="0" value={editEmp.tunjMakan} onChange={e=>setEditEmp(p=>({...p,tunjMakan:Number(e.target.value)}))} style={inp}/></F>
              <F label="T. Transport (IDR)"><input type="number" min="0" value={editEmp.tunjTransport} onChange={e=>setEditEmp(p=>({...p,tunjTransport:Number(e.target.value)}))} style={inp}/></F>
              <F label="T. Jabatan (IDR)"><input type="number" min="0" value={editEmp.tunjJabatan} onChange={e=>setEditEmp(p=>({...p,tunjJabatan:Number(e.target.value)}))} style={inp}/></F>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginTop:10}}>
              <F label="Pot. BPJS (IDR)"><input type="number" min="0" value={editEmp.potonganBPJS} onChange={e=>setEditEmp(p=>({...p,potonganBPJS:Number(e.target.value)}))} style={inp}/></F>
              <F label="Pot. PPh21 (IDR)"><input type="number" min="0" value={editEmp.potonganPPh} onChange={e=>setEditEmp(p=>({...p,potonganPPh:Number(e.target.value)}))} style={inp}/></F>
            </div>
            {editEmp.gajiPokok>0&&<div style={{marginTop:10,background:"#FFF",borderRadius:8,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",border:`1px solid ${C.border}`}}>
              <span style={{fontSize:13,color:C.muted}}>Take Home Estimasi</span>
              <span style={{fontWeight:800,color:C.primary,fontSize:16}}>{fmt(editEmp.gajiPokok+editEmp.tunjMakan+editEmp.tunjTransport+editEmp.tunjJabatan-editEmp.potonganBPJS-editEmp.potonganPPh)}</span>
            </div>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Bank"><input value={editEmp.bank} onChange={e=>setEditEmp(p=>({...p,bank:e.target.value}))} style={inp} placeholder="BCA, Mandiri..."/></F>
            <F label="No. Rekening"><input value={editEmp.noRek} onChange={e=>setEditEmp(p=>({...p,noRek:e.target.value}))} style={inp}/></F>
          </div>
          <button onClick={()=>saveEmployee(editEmp)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* Edit Payroll Modal */}
      {editPay&&<Modal onClose={()=>setEditPay(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>✏️ Edit Data Penggajian</div>
        <div style={{display:"grid",gap:12}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Lembur (IDR)"><input type="number" value={editPay.lemburAmount} onChange={e=>setEditPay(p=>{const lA=Number(e.target.value);const tT=p.tunjMakan+p.tunjTransport+p.tunjJabatan+lA;const tP=p.potonganBPJS+p.potonganPPh+p.potonganLain;return{...p,lemburAmount:lA,totalTunjangan:tT,totalPotongan:tP,takehome:p.gajiPokok+tT-tP};})  } style={inp}/></F>
            <F label="Potongan Lain (IDR)"><input type="number" value={editPay.potonganLain} onChange={e=>setEditPay(p=>{const pL=Number(e.target.value);const tP=p.potonganBPJS+p.potonganPPh+pL;return{...p,potonganLain:pL,totalPotongan:tP,takehome:p.gajiPokok+p.totalTunjangan-tP};})} style={inp}/></F>
          </div>
          <F label="Status"><select value={editPay.status} onChange={e=>setEditPay(p=>({...p,status:e.target.value,payDate:e.target.value==="dibayar"?(p.payDate||today()):null}))} style={sel}><option value="proses">Diproses</option><option value="dibayar">Dibayar</option></select></F>
          {editPay.status==="dibayar"&&<F label="Tanggal Bayar"><input type="date" value={editPay.payDate||""} onChange={e=>setEditPay(p=>({...p,payDate:e.target.value}))} style={inp}/></F>}
          <div style={{background:C.surface,borderRadius:8,padding:12,fontSize:13}}>
            <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:C.muted}}>Take Home</span><span style={{fontWeight:800,color:C.primary,fontSize:15}}>{fmt(editPay.takehome)}</span></div>
          </div>
          <button onClick={()=>savePayroll(editPay)} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* Slip Gaji Modal */}
      {slipModal&&<Modal onClose={()=>setSlipModal(null)}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontWeight:700,fontSize:17}}>📄 Slip Gaji</div>
          <button onClick={()=>{
            const html=`<h2>Slip Gaji — ${slipModal.employeeName}</h2>
            <div style="display:flex;justify-content:space-between;background:#F8FAFC;padding:14px 18px;border-radius:8px;margin-bottom:16px">
              <div><b>${slipModal.employeeName}</b><br/><span style="color:#64748B">${employees.find(e=>e.id===slipModal.employeeId)?.position||""}</span></div>
              <div style="text-align:right"><b>${slipModal.month}</b><br/><span style="color:#64748B">ID: ${slipModal.id}</span></div>
            </div>
            <table><thead><tr><th>Komponen</th><th>Keterangan</th><th style="text-align:right">Jumlah</th></tr></thead><tbody>
              <tr><td><b>Gaji Pokok</b></td><td></td><td style="text-align:right">${fmt(slipModal.gajiPokok)}</td></tr>
              <tr style="color:#10B97B"><td>Tunjangan Makan</td><td></td><td style="text-align:right">+${fmt(slipModal.tunjMakan)}</td></tr>
              <tr style="color:#10B97B"><td>Tunjangan Transport</td><td></td><td style="text-align:right">+${fmt(slipModal.tunjTransport)}</td></tr>
              ${slipModal.tunjJabatan>0?`<tr style="color:#10B97B"><td>Tunjangan Jabatan</td><td></td><td style="text-align:right">+${fmt(slipModal.tunjJabatan)}</td></tr>`:""}
              ${slipModal.lemburAmount>0?`<tr style="color:#10B97B"><td>Lembur</td><td>${slipModal.lemburJam} jam</td><td style="text-align:right">+${fmt(slipModal.lemburAmount)}</td></tr>`:""}
              <tr style="color:#EF4444"><td>Potongan BPJS</td><td></td><td style="text-align:right">-${fmt(slipModal.potonganBPJS)}</td></tr>
              <tr style="color:#EF4444"><td>Potongan PPh21</td><td></td><td style="text-align:right">-${fmt(slipModal.potonganPPh)}</td></tr>
              ${slipModal.potonganLain>0?`<tr style="color:#EF4444"><td>Potongan Lain</td><td>${slipModal.potonganLainKet||""}</td><td style="text-align:right">-${fmt(slipModal.potonganLain)}</td></tr>`:""}
            </tbody></table>
            <div style="background:#1B5FBF;color:#FFF;padding:14px 18px;border-radius:8px;display:flex;justify-content:space-between;align-items:center;font-size:15px">
              <b>TAKE HOME PAY</b><b style="font-size:20px">${fmt(slipModal.takehome)}</b>
            </div>
            <div style="margin-top:16px;display:flex;justify-content:space-between;font-size:12px;color:#64748B">
              <span>Tanggal Bayar: ${slipModal.payDate||"Belum dibayar"}</span><span>Bank: ${employees.find(e=>e.id===slipModal.employeeId)?.bank||"-"} – ${employees.find(e=>e.id===slipModal.employeeId)?.noRek||"-"}</span>
            </div>`;
            exportPDF(`Slip Gaji ${slipModal.employeeName} ${slipModal.month}`,html);
          }} style={btn(C.primary)}>🖨️ Cetak Slip</button>
        </div>
        <div style={{background:C.surface,borderRadius:10,padding:16}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
            <div><div style={{fontWeight:700,fontSize:15}}>{slipModal.employeeName}</div><div style={{fontSize:12,color:C.muted}}>{employees.find(e=>e.id===slipModal.employeeId)?.position} · {slipModal.month}</div></div>
            <div style={{textAlign:"right"}}><Badge status={slipModal.status}/><div style={{fontSize:11,color:C.muted,marginTop:4}}>{slipModal.id}</div></div>
          </div>
          {[
            ["Gaji Pokok",slipModal.gajiPokok,C.text],
            ["Tunjangan Makan",slipModal.tunjMakan,C.success],
            ["Tunjangan Transport",slipModal.tunjTransport,C.success],
            ...(slipModal.tunjJabatan>0?[["Tunjangan Jabatan",slipModal.tunjJabatan,C.success]]:[]),
            ...(slipModal.lemburAmount>0?[["Lembur ("+slipModal.lemburJam+" jam)",slipModal.lemburAmount,C.success]]:[]),
            null,
            ["Pot. BPJS",-slipModal.potonganBPJS,C.danger],
            ["Pot. PPh21",-slipModal.potonganPPh,C.danger],
            ...(slipModal.potonganLain>0?[["Pot. Lain ("+slipModal.potonganLainKet+")",-slipModal.potonganLain,C.danger]]:[]),
          ].map((row,i)=>row===null?<div key={i} style={{height:1,background:C.border,margin:"8px 0"}}/>:(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",fontSize:13}}>
              <span style={{color:C.muted}}>{row[0]}</span>
              <span style={{fontWeight:500,color:row[2]}}>{row[1]<0?"-":"++"[0]}{fmt(Math.abs(row[1]))}</span>
            </div>
          ))}
          <div style={{marginTop:12,background:C.primary,borderRadius:8,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:700,color:"#FFF"}}>Take Home Pay</span>
            <span style={{fontWeight:800,fontSize:18,color:"#FFF"}}>{fmt(slipModal.takehome)}</span>
          </div>
          {slipModal.payDate&&<div style={{marginTop:10,fontSize:12,color:C.muted,textAlign:"right"}}>Dibayar: {slipModal.payDate} · {employees.find(e=>e.id===slipModal.employeeId)?.bank} {employees.find(e=>e.id===slipModal.employeeId)?.noRek}</div>}
        </div>
      </Modal>}

      {confirmDel&&<ConfirmModal msg={`Hapus karyawan "${employees.find(e=>e.id===confirmDel)?.name}"? Semua data penggajian terkait juga dihapus.`} onConfirm={()=>deleteEmployee(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
      {confirmDelPay&&<ConfirmModal msg="Hapus data penggajian ini?" onConfirm={()=>deletePayroll(confirmDelPay)} onCancel={()=>setConfirmDelPay(null)}/>}
    </div>
  );
}

// ─── LAPORAN PAGE ─────────────────────────────────────────────────────────────
function LaporanPage({products,movements,orders,invoices,cashFlowEntries,employees,payrolls}){
  const [period,setPeriod]=useState("2024-06");

  const ordersP=orders.filter(o=>!period||o.date.startsWith(period));
  const invoicesP=invoices.filter(i=>!period||i.date.startsWith(period));
  const paidInv=invoicesP.filter(i=>i.status==="paid");
  const cashIn=cashFlowEntries.filter(e=>e.type==="in"&&(!period||e.date.startsWith(period))).reduce((s,e)=>s+e.amount,0);
  const cashOut=cashFlowEntries.filter(e=>e.type==="out"&&(!period||e.date.startsWith(period))).reduce((s,e)=>s+e.amount,0);
  const payrollsP=payrolls.filter(p=>!period||p.month===period);
  const totalGaji=payrollsP.reduce((s,p)=>s+p.takehome,0);
  const nilaiInventori=products.reduce((s,p)=>s+p.stock*p.price,0);
  const pendapatan=paidInv.reduce((s,i)=>s+i.grandTotal,0);
  const labaKotor=pendapatan-totalGaji-(cashOut-totalGaji>0?cashOut-totalGaji:0);

  const kpiData=[
    {l:"Pendapatan (Invoice Lunas)",v:fmt(pendapatan),c:C.success,i:"💰",sub:`${paidInv.length} invoice`},
    {l:"Total Pesanan",v:ordersP.length,c:C.primary,i:"🛒",sub:`Nilai: ${fmt(ordersP.reduce((s,o)=>s+o.total,0))}`},
    {l:"Total Pengeluaran Kas",v:fmt(cashOut),c:C.danger,i:"📤",sub:`Pemasukan: ${fmt(cashIn)}`},
    {l:"Gaji Karyawan",v:fmt(totalGaji),c:C.warning,i:"👥",sub:`${payrollsP.length} karyawan`},
    {l:"Arus Kas Bersih",v:fmt(cashIn-cashOut),c:(cashIn-cashOut)>=0?C.success:C.danger,i:(cashIn-cashOut)>=0?"📈":"📉",sub:(cashIn-cashOut)>=0?"Surplus":"Defisit"},
    {l:"Nilai Inventori",v:fmt(nilaiInventori),c:C.purple,i:"📦",sub:`${products.length} SKU`},
  ];

  // Departemen breakdown
  const deptGaji=payrolls.reduce((acc,p)=>{const emp=employees.find(e=>e.id===p.employeeId);if(emp){acc[emp.dept]=(acc[emp.dept]||0)+p.takehome;}return acc;},{});

  // Top produk by nilai
  const topProd=[...products].sort((a,b)=>(b.stock*b.price)-(a.stock*a.price)).slice(0,5);

  // Order status breakdown
  const statusCount=orders.reduce((acc,o)=>{acc[o.status]=(acc[o.status]||0)+1;return acc;},{});

  const handleFullReport=()=>{
    const invHtml=buildInvExport(products).html;
    const ordHtml=buildOrdExport(orders).html;
    const invHtml2=buildInvoExport(invoices).html;
    const cfHtml=buildCFExport(cashFlowEntries).html;
    const empHtml=buildEmpExport(employees).html;
    const payHtml=buildPayrollExport(payrolls).html;
    const summary=`
      <div class="sum">
        <div><label>Pendapatan</label><span style="color:#10B97B">${fmt(pendapatan)}</span></div>
        <div><label>Total Pengeluaran</label><span style="color:#EF4444">${fmt(cashOut)}</span></div>
        <div><label>Gaji Karyawan</label><span style="color:#F59E0B">${fmt(totalGaji)}</span></div>
        <div><label>Arus Kas Bersih</label><span style="color:${cashIn-cashOut>=0?"#10B97B":"#EF4444"}">${fmt(cashIn-cashOut)}</span></div>
        <div><label>Nilai Inventori</label><span>${fmt(nilaiInventori)}</span></div>
      </div>`;
    exportPDF(`Laporan Lengkap ERP ${period||"Semua"}`,summary+invHtml+ordHtml+invHtml2+cfHtml+empHtml+payHtml);
  };

  const handleFullExcel=()=>{
    exportExcel([
      {name:"Inventori",...buildInvExport(products)},
      {name:"Mutasi Stok",...buildMovExport(movements)},
      {name:"Pesanan",...buildOrdExport(orders)},
      {name:"Invoice",...buildInvoExport(invoices)},
      {name:"Arus Kas",...buildCFExport(cashFlowEntries)},
      {name:"Karyawan",...buildEmpExport(employees)},
      {name:"Penggajian",...buildPayrollExport(payrolls)},
    ]);
  };

  return(
    <div>
      {/* Header + period filter */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <div style={{fontWeight:700,fontSize:16}}>Laporan Bisnis</div>
          <div style={{fontSize:12,color:C.muted}}>Ringkasan performa semua modul</div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <input type="month" value={period} onChange={e=>setPeriod(e.target.value)} style={{...inp,width:160}}/>
          <button onClick={()=>setPeriod("")} style={{...btn("#94A3B8"),padding:"8px 12px",fontSize:12}}>Semua</button>
          <button onClick={handleFullExcel} style={{...btn(C.success),display:"flex",alignItems:"center",gap:6}}>📊 Excel Lengkap</button>
          <button onClick={handleFullReport} style={{...btn(C.primary),display:"flex",alignItems:"center",gap:6}}>📄 PDF Lengkap</button>
        </div>
      </div>

      {/* KPI Grid */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:20}}>
        {kpiData.map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div><div style={{fontSize:12,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:20,fontWeight:700,color:k.c}}>{k.v}</div><div style={{fontSize:11,color:C.muted,marginTop:3}}>{k.sub}</div></div>
              <span style={{fontSize:26}}>{k.i}</span>
            </div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        {/* Laporan Keuangan */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>📊 Ringkasan Keuangan {period}</div>
          {[
            {l:"Total Pendapatan (Invoice Lunas)",v:pendapatan,c:C.success,sign:"+"},
            {l:"Pengeluaran Operasional",v:cashOut,c:C.danger,sign:"-"},
            {l:"Total Gaji Karyawan",v:totalGaji,c:C.warning,sign:"-"},
          ].map((row,i)=>(
            <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"10px 0",borderBottom:`1px solid ${C.border}`}}>
              <span style={{fontSize:13,color:C.muted}}>{row.l}</span>
              <span style={{fontWeight:700,color:row.c}}>{row.sign}{fmt(row.v)}</span>
            </div>
          ))}
          <div style={{display:"flex",justifyContent:"space-between",padding:"12px 0",marginTop:4}}>
            <span style={{fontWeight:700,fontSize:14}}>Arus Kas Bersih</span>
            <span style={{fontWeight:800,fontSize:17,color:(cashIn-cashOut)>=0?C.success:C.danger}}>{fmt(cashIn-cashOut)}</span>
          </div>
          <div style={{marginTop:8}}>
            <ExportBtn onExcel={()=>exportExcel([{name:"Keuangan",...buildCFExport(cashFlowEntries)}])} onPdf={()=>exportPDF("Laporan Keuangan",buildCFExport(cashFlowEntries).html+buildInvoExport(invoices).html)}/>
          </div>
        </div>

        {/* Laporan Gaji per Departemen */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>👥 Gaji per Departemen</div>
          {Object.entries(deptGaji).sort((a,b)=>b[1]-a[1]).map(([dept,total])=>{
            const pct=Math.round((total/Object.values(deptGaji).reduce((s,v)=>s+v,0))*100);
            return(
              <div key={dept} style={{marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:5}}>
                  <span style={{fontWeight:600}}>{dept}</span>
                  <span style={{color:C.muted}}>{pct}% · {fmt(total)}</span>
                </div>
                <div style={{background:C.surface,borderRadius:5,height:8}}>
                  <div style={{width:`${pct}%`,height:"100%",background:C.warning,borderRadius:5}}/>
                </div>
              </div>
            );
          })}
          <div style={{marginTop:8}}>
            <ExportBtn onExcel={()=>exportExcel([{name:"Penggajian",...buildPayrollExport(payrolls)}])} onPdf={()=>exportPDF("Laporan Penggajian",buildPayrollExport(payrolls).html+buildEmpExport(employees).html)}/>
          </div>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        {/* Status Pesanan */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>🛒 Status Pesanan</div>
          {Object.entries(statusCount).map(([st,cnt])=>{
            const total=orders.length;
            const pct=Math.round((cnt/total)*100);
            const colors={pending:C.warning,processing:C.primary,delivered:C.success,cancelled:C.danger};
            return(
              <div key={st} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>
                <Badge status={st}/>
                <div style={{flex:1}}>
                  <div style={{background:C.surface,borderRadius:4,height:6}}><div style={{width:`${pct}%`,height:"100%",background:colors[st]||C.muted,borderRadius:4}}/></div>
                </div>
                <span style={{fontWeight:700,fontSize:14}}>{cnt}</span>
                <span style={{fontSize:12,color:C.muted}}>{pct}%</span>
              </div>
            );
          })}
          <div style={{marginTop:12}}>
            <ExportBtn onExcel={()=>exportExcel([{name:"Pesanan",...buildOrdExport(orders)}])} onPdf={()=>exportPDF("Laporan Pesanan",buildOrdExport(orders).html)}/>
          </div>
        </div>

        {/* Top produk nilai stok */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>📦 Top 5 Produk (Nilai Stok)</div>
          {topProd.map((p,i)=>(
            <div key={p.id} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>
              <div style={{width:24,height:24,borderRadius:6,background:C.primaryL,color:C.primary,fontWeight:700,fontSize:12,display:"flex",alignItems:"center",justifyContent:"center"}}>{i+1}</div>
              <div style={{flex:1}}><div style={{fontWeight:600,fontSize:13}}>{p.name}</div><div style={{fontSize:11,color:C.muted}}>{p.stock} unit × {fmt(p.price)}</div></div>
              <span style={{fontWeight:700,color:p.stock<=p.minStock?C.danger:C.text}}>{fmt(p.stock*p.price)}</span>
            </div>
          ))}
          <div style={{marginTop:12}}>
            <ExportBtn onExcel={()=>exportExcel([{name:"Inventori",...buildInvExport(products)}])} onPdf={()=>exportPDF("Laporan Inventori",buildInvExport(products).html)}/>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── PURCHASE ORDER PAGE ──────────────────────────────────────────────────────
function POPage({pos,setPOs,suppliers,setSuppliers,products,setProducts,setMovements}){
  const [tab,setTab]=useState("daftar");
  const [editPO,setEditPO]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [editSupplier,setEditSupplier]=useState(null);
  const [confirmDelSup,setConfirmDelSup]=useState(null);
  const [viewPO,setViewPO]=useState(null);
  const [poItems,setPOItems]=useState([{productId:"",name:"",qty:1,price:0}]);

  const savePO=(po)=>{
    if(po._new){setPOs(prev=>[...prev,{...po,id:`PO-${Date.now()}`,_new:undefined}]);}
    else{setPOs(prev=>prev.map(x=>x.id===po.id?po:x));}
    setEditPO(null);
  };
  const deletePO=(id)=>{setPOs(prev=>prev.filter(p=>p.id!==id));setConfirmDel(null);};

  const terimaPO=(po)=>{
    // update stock
    po.items.forEach(item=>{
      const prod=products.find(p=>p.id===item.productId);
      if(!prod)return;
      const before=prod.stock, after=before+item.qty;
      setProducts(prev=>prev.map(p=>p.id===item.productId?{...p,stock:after}:p));
      setMovements(prev=>[...prev,{id:Date.now()+Math.random(),date:today(),type:"in",subtype:"pembelian",product:item.name,productId:item.productId,qty:item.qty,ref:po.id,note:`Penerimaan PO dari ${po.supplierName}`,before,after}]);
    });
    setPOs(prev=>prev.map(p=>p.id===po.id?{...p,status:"diterima",receivedDate:today()}:p));
  };

  const saveSupplier=(s)=>{
    if(s._new){setSuppliers(prev=>[...prev,{...s,id:Date.now(),_new:undefined}]);}
    else{setSuppliers(prev=>prev.map(x=>x.id===s.id?s:x));}
    setEditSupplier(null);
  };
  const deleteSupplier=(id)=>{setSuppliers(prev=>prev.filter(s=>s.id!==id));setConfirmDelSup(null);};

  const addItem=()=>setPOItems(p=>[...p,{productId:"",name:"",qty:1,price:0}]);
  const updateItem=(i,field,val)=>setPOItems(p=>p.map((item,idx)=>{
    if(idx!==i)return item;
    if(field==="productId"){const prod=products.find(x=>x.id===Number(val));return{...item,productId:Number(val),name:prod?.name||"",price:prod?Math.round(prod.price*0.7):0};}
    return{...item,[field]:field==="qty"||field==="price"?Number(val):val};
  }));
  const removeItem=(i)=>setPOItems(p=>p.filter((_,idx)=>idx!==i));

  const totalPO=poItems.reduce((s,it)=>s+it.qty*it.price,0);

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total PO",v:pos.length,c:C.primary,i:"📋"},
          {l:"Menunggu Terima",v:pos.filter(p=>p.status==="dipesan").length,c:C.warning,i:"⏳"},
          {l:"Sudah Diterima",v:pos.filter(p=>p.status==="diterima").length,c:C.success,i:"✅"},
          {l:"Total Nilai PO",v:fmt(pos.reduce((s,p)=>s+p.total,0)),c:C.purple,i:"💰"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div></div><span style={{fontSize:24}}>{k.i}</span></div>
          </div>
        ))}
      </div>

      <div style={{display:"flex",gap:4,marginBottom:16,background:"#FFF",borderRadius:12,padding:6,border:`1px solid ${C.border}`,width:"fit-content",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        {[{id:"daftar",l:"📋 Daftar PO"},{id:"buat",l:"➕ Buat PO"},{id:"supplier",l:"🏭 Supplier"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?C.primary:"transparent",color:tab===t.id?"#FFF":C.muted}}>{t.l}</button>
        ))}
      </div>

      {/* DAFTAR PO */}
      {tab==="daftar"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:15}}>Daftar Purchase Order</div>
            <ExportBtn onExcel={()=>exportExcel([{name:"PO",...buildPOExport(pos)}])} onPdf={()=>exportPDF("Laporan Purchase Order",buildPOExport(pos).html)}/>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["No PO","Tanggal","Supplier","Items","Subtotal","PPN","Total","Status","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {pos.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada Purchase Order</td></tr>
                :pos.slice().reverse().map(p=>(
                <tr key={p.id} style={{borderBottom:`1px solid ${C.border}`}}>
                  <td style={tdS}><span style={{color:C.primary,fontWeight:700}}>{p.id}</span></td>
                  <td style={tdS}>{p.date}</td>
                  <td style={tdS}><div style={{fontWeight:600}}>{p.supplierName}</div></td>
                  <td style={tdS}><span style={{color:C.muted,fontSize:12}}>{p.items.length} item · {p.items.reduce((s,it)=>s+it.qty,0)} unit</span></td>
                  <td style={tdS}>{fmt(p.subtotal)}</td>
                  <td style={tdS}>{fmt(p.tax)}</td>
                  <td style={tdS}><span style={{fontWeight:700}}>{fmt(p.total)}</span></td>
                  <td style={tdS}><Badge status={p.status}/></td>
                  <td style={tdS}>
                    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                      <button onClick={()=>setViewPO(p)} style={btnSm(C.primary)}>📄 Detail</button>
                      {p.status==="dipesan"&&<button onClick={()=>terimaPO(p)} style={btnSm(C.success)}>✓ Terima</button>}
                      <button onClick={()=>setEditPO({...p})} style={btnSm("#64748B")}>✏️</button>
                      <button onClick={()=>setConfirmDel(p.id)} style={btnSm(C.danger)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* BUAT PO */}
      {tab==="buat"&&(
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:18,color:C.primary}}>➕ Buat Purchase Order Baru</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:16}}>
            <F label="Supplier *">
              <select style={sel} id="po-sup" onChange={()=>{}}>
                <option value="">— Pilih supplier —</option>
                {suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </F>
            <F label="Tanggal PO"><input type="date" defaultValue={today()} style={inp} id="po-date"/></F>
            <F label="Catatan"><input type="text" placeholder="Keterangan..." style={inp} id="po-notes"/></F>
          </div>
          <div style={{fontWeight:600,fontSize:13,marginBottom:10}}>Item Pembelian</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,marginBottom:12}}>
            <thead><tr style={{background:C.surface}}>{["Produk","Qty","Harga Beli (IDR)","Subtotal",""].map(h=><th key={h} style={{...thS,borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
            <tbody>
              {poItems.map((item,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${C.border}`}}>
                  <td style={tdS}>
                    <select value={item.productId} onChange={e=>updateItem(i,"productId",e.target.value)} style={{...sel,width:"100%"}}>
                      <option value="">— Pilih produk —</option>
                      {products.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </td>
                  <td style={tdS}><input type="number" min="1" value={item.qty} onChange={e=>updateItem(i,"qty",e.target.value)} style={{...inp,width:80,textAlign:"center"}}/></td>
                  <td style={tdS}><input type="number" min="0" value={item.price} onChange={e=>updateItem(i,"price",e.target.value)} style={{...inp,width:150}}/></td>
                  <td style={tdS}><span style={{fontWeight:700}}>{fmt(item.qty*item.price)}</span></td>
                  <td style={tdS}><button onClick={()=>removeItem(i)} style={{...btnSm(C.danger),padding:"3px 8px"}}>✕</button></td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
            <button onClick={addItem} style={{...btn("#64748B"),padding:"7px 14px",fontSize:12}}>+ Tambah Item</button>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:13,color:C.muted}}>Subtotal: <b>{fmt(totalPO)}</b></div>
              <div style={{fontSize:13,color:C.muted}}>PPN 10%: <b>{fmt(totalPO*0.1)}</b></div>
              <div style={{fontWeight:700,fontSize:16,color:C.primary}}>Total: {fmt(totalPO*1.1)}</div>
            </div>
          </div>
          <button onClick={()=>{
            const supEl=document.getElementById("po-sup");
            const sup=suppliers.find(s=>s.id===Number(supEl?.value));
            if(!sup||poItems.every(it=>!it.productId)){alert("Pilih supplier dan minimal 1 produk!");return;}
            const items=poItems.filter(it=>it.productId);
            const sub=items.reduce((s,it)=>s+it.qty*it.price,0);
            const newPO={id:`PO-${Date.now()}`,date:document.getElementById("po-date")?.value||today(),supplierId:sup.id,supplierName:sup.name,status:"dipesan",items,subtotal:sub,tax:sub*0.1,total:sub*1.1,notes:document.getElementById("po-notes")?.value||"",receivedDate:null};
            setPOs(prev=>[...prev,newPO]);
            setPOItems([{productId:"",name:"",qty:1,price:0}]);
            setTab("daftar");
          }} style={{...btn(C.primary),padding:"11px 24px"}}>📋 Simpan Purchase Order</button>
        </div>
      )}

      {/* SUPPLIER */}
      {tab==="supplier"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:15}}>Data Supplier</div>
            <button onClick={()=>setEditSupplier({_new:true,name:"",contact:"",phone:"",email:"",address:"",terms:"30 hari"})} style={btn(C.primary)}>+ Tambah Supplier</button>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["Nama","Kontak","Telepon","Email","Alamat","Terms","Total PO","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {suppliers.map(s=>{
                const spos=pos.filter(p=>p.supplierId===s.id);
                return(
                  <tr key={s.id} style={{borderBottom:`1px solid ${C.border}`}}>
                    <td style={tdS}><div style={{fontWeight:700}}>{s.name}</div></td>
                    <td style={tdS}>{s.contact}</td>
                    <td style={tdS}>{s.phone}</td>
                    <td style={tdS}><a href={`mailto:${s.email}`} style={{color:C.primary,fontSize:12}}>{s.email}</a></td>
                    <td style={tdS}><div style={{fontSize:12,color:C.muted,maxWidth:160}}>{s.address}</div></td>
                    <td style={tdS}><span style={{background:C.primaryL,color:C.primary,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:500}}>{s.terms}</span></td>
                    <td style={tdS}><div style={{fontWeight:700,color:C.primary}}>{fmt(spos.reduce((s,p)=>s+p.total,0))}</div><div style={{fontSize:11,color:C.muted}}>{spos.length} PO</div></td>
                    <td style={tdS}><div style={{display:"flex",gap:5}}><button onClick={()=>setEditSupplier({...s})} style={btnSm("#64748B")}>✏️</button><button onClick={()=>setConfirmDelSup(s.id)} style={btnSm(C.danger)}>🗑️</button></div></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Detail PO Modal */}
      {viewPO&&<Modal onClose={()=>setViewPO(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:4}}>Detail Purchase Order</div>
        <div style={{color:C.primary,fontWeight:600,marginBottom:14}}>{viewPO.id}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
          {[["Supplier",viewPO.supplierName],["Tanggal",viewPO.date],["Status",<Badge status={viewPO.status}/>],["Tgl Terima",viewPO.receivedDate||"—"]].map(([k,v])=>(
            <div key={k} style={{background:C.surface,padding:10,borderRadius:8}}><div style={{fontSize:11,color:C.muted,marginBottom:3}}>{k}</div><div style={{fontWeight:600}}>{v}</div></div>
          ))}
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,marginBottom:14}}>
          <thead><tr style={{background:C.surface}}>{["Produk","Qty","Harga Beli","Subtotal"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>{viewPO.items.map((it,i)=><tr key={i} style={{borderBottom:`1px solid ${C.border}`}}><td style={tdS}>{it.name}</td><td style={tdS}>{it.qty}</td><td style={tdS}>{fmt(it.price)}</td><td style={tdS}><b>{fmt(it.qty*it.price)}</b></td></tr>)}</tbody>
        </table>
        <div style={{textAlign:"right",borderTop:`2px solid ${C.border}`,paddingTop:10}}>
          <div style={{fontSize:13,color:C.muted}}>Subtotal: {fmt(viewPO.subtotal)}</div>
          <div style={{fontSize:13,color:C.muted}}>PPN 10%: {fmt(viewPO.tax)}</div>
          <div style={{fontWeight:700,fontSize:16,color:C.primary,marginTop:4}}>Total: {fmt(viewPO.total)}</div>
        </div>
        {viewPO.status==="dipesan"&&<button onClick={()=>{terimaPO(viewPO);setViewPO(null);}} style={{...btn(C.success),marginTop:14,width:"100%"}}>✓ Terima Barang & Update Stok</button>}
      </Modal>}

      {/* Edit PO */}
      {editPO&&<Modal onClose={()=>setEditPO(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:16}}>✏️ Edit Purchase Order</div>
        <div style={{display:"grid",gap:12}}>
          <F label="Status"><select value={editPO.status} onChange={e=>setEditPO(p=>({...p,status:e.target.value}))} style={sel}>{["dipesan","diterima","sebagian","dibatalkan"].map(s=><option key={s}>{s}</option>)}</select></F>
          <F label="Tanggal Terima"><input type="date" value={editPO.receivedDate||""} onChange={e=>setEditPO(p=>({...p,receivedDate:e.target.value}))} style={inp}/></F>
          <F label="Catatan"><input value={editPO.notes||""} onChange={e=>setEditPO(p=>({...p,notes:e.target.value}))} style={inp}/></F>
          <button onClick={()=>savePO(editPO)} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* Edit Supplier */}
      {editSupplier&&<Modal onClose={()=>setEditSupplier(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:16}}>{editSupplier._new?"➕ Tambah Supplier":"✏️ Edit Supplier"}</div>
        <div style={{display:"grid",gap:12}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Nama Perusahaan *"><input value={editSupplier.name} onChange={e=>setEditSupplier(p=>({...p,name:e.target.value}))} style={inp}/></F>
            <F label="Nama Kontak"><input value={editSupplier.contact} onChange={e=>setEditSupplier(p=>({...p,contact:e.target.value}))} style={inp}/></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Telepon"><input value={editSupplier.phone} onChange={e=>setEditSupplier(p=>({...p,phone:e.target.value}))} style={inp}/></F>
            <F label="Email"><input type="email" value={editSupplier.email} onChange={e=>setEditSupplier(p=>({...p,email:e.target.value}))} style={inp}/></F>
          </div>
          <F label="Alamat"><input value={editSupplier.address} onChange={e=>setEditSupplier(p=>({...p,address:e.target.value}))} style={inp}/></F>
          <F label="Terms Pembayaran"><input value={editSupplier.terms} onChange={e=>setEditSupplier(p=>({...p,terms:e.target.value}))} style={inp} placeholder="30 hari, COD..."/></F>
          <button onClick={()=>saveSupplier(editSupplier)} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}

      {confirmDel&&<ConfirmModal msg="Hapus PO ini?" onConfirm={()=>deletePO(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
      {confirmDelSup&&<ConfirmModal msg={`Hapus supplier "${suppliers.find(s=>s.id===confirmDelSup)?.name}"?`} onConfirm={()=>deleteSupplier(confirmDelSup)} onCancel={()=>setConfirmDelSup(null)}/>}
    </div>
  );
}

// ─── RETUR PAGE ───────────────────────────────────────────────────────────────
function ReturPage({returns,setReturns,products,setProducts,setMovements,orders,customers,pos,suppliers}){
  const [editRetur,setEditRetur]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [form,setForm]=useState({type:"penjualan",refId:"",customer:"",supplier:"",date:today(),reason:"",items:[{productId:"",name:"",qty:1,price:0}],notes:""});

  const addItem=()=>setForm(f=>({...f,items:[...f.items,{productId:"",name:"",qty:1,price:0}]}));
  const updateItem=(i,field,val)=>setForm(f=>({...f,items:f.items.map((it,idx)=>{
    if(idx!==i)return it;
    if(field==="productId"){const prod=products.find(x=>x.id===Number(val));return{...it,productId:Number(val),name:prod?.name||"",price:prod?.price||0};}
    return{...it,[field]:field==="qty"||field==="price"?Number(val):val};
  })}));

  const submitRetur=()=>{
    const items=form.items.filter(it=>it.productId&&it.qty>0);
    if(!items.length||!form.reason){alert("Lengkapi data retur!");return;}
    const total=items.reduce((s,it)=>s+it.qty*it.price,0);
    const newRetur={id:`RTN-${Date.now()}`,date:form.date,type:form.type,refId:form.refId,customer:form.type==="penjualan"?form.customer:"",supplier:form.type==="pembelian"?form.supplier:"",items,total,reason:form.reason,status:"menunggu",notes:form.notes};
    setReturns(prev=>[...prev,newRetur]);
    setForm({type:"penjualan",refId:"",customer:"",supplier:"",date:today(),reason:"",items:[{productId:"",name:"",qty:1,price:0}],notes:""});
  };

  const setujuiRetur=(retur)=>{
    // retur penjualan → stok masuk kembali; retur pembelian → stok keluar
    retur.items.forEach(item=>{
      const prod=products.find(p=>p.id===item.productId);
      if(!prod)return;
      const before=prod.stock;
      const delta=retur.type==="penjualan"?item.qty:-item.qty;
      const after=before+delta;
      setProducts(prev=>prev.map(p=>p.id===item.productId?{...p,stock:Math.max(0,after)}:p));
      setMovements(prev=>[...prev,{id:Date.now()+Math.random(),date:today(),type:delta>0?"in":"out",subtype:"retur",product:item.name,productId:item.productId,qty:item.qty,ref:retur.id,note:`Retur ${retur.type}: ${retur.reason}`,before,after:Math.max(0,after)}]);
    });
    setReturns(prev=>prev.map(r=>r.id===retur.id?{...r,status:"disetujui"}:r));
  };
  const tolakRetur=(id)=>setReturns(prev=>prev.map(r=>r.id===id?{...r,status:"ditolak"}:r));
  const deleteRetur=(id)=>{setReturns(prev=>prev.filter(r=>r.id!==id));setConfirmDel(null);};

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total Retur",v:returns.length,c:C.primary,i:"↩️"},
          {l:"Menunggu",v:returns.filter(r=>r.status==="menunggu").length,c:C.warning,i:"⏳"},
          {l:"Disetujui",v:returns.filter(r=>r.status==="disetujui").length,c:C.success,i:"✅"},
          {l:"Total Nilai Retur",v:fmt(returns.reduce((s,r)=>s+r.total,0)),c:C.danger,i:"💸"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div></div><span style={{fontSize:24}}>{k.i}</span></div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"380px 1fr",gap:16,alignItems:"start"}}>
        {/* Form Retur */}
        <div style={{...card,borderTop:`4px solid ${C.danger}`}}>
          <div style={{fontWeight:700,fontSize:15,color:C.danger,marginBottom:4}}>↩️ Buat Retur Baru</div>
          <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Catat retur penjualan atau pembelian</div>
          <div style={{display:"grid",gap:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <F label="Jenis Retur">
                <select value={form.type} onChange={e=>setForm(f=>({...f,type:e.target.value}))} style={sel}>
                  <option value="penjualan">Retur Penjualan</option>
                  <option value="pembelian">Retur Pembelian</option>
                </select>
              </F>
              <F label="Tanggal"><input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))} style={inp}/></F>
            </div>
            {form.type==="penjualan"
              ?<F label="Pelanggan"><select value={form.customer} onChange={e=>setForm(f=>({...f,customer:e.target.value}))} style={sel}><option value="">— Pilih pelanggan —</option>{customers.map(c=><option key={c.id}>{c.name}</option>)}</select></F>
              :<F label="Supplier"><select value={form.supplier} onChange={e=>setForm(f=>({...f,supplier:e.target.value}))} style={sel}><option value="">— Pilih supplier —</option>{suppliers.map(s=><option key={s.id}>{s.name}</option>)}</select></F>
            }
            <F label="No. Referensi (SO/PO)"><input value={form.refId} onChange={e=>setForm(f=>({...f,refId:e.target.value}))} style={inp} placeholder="SO-xxx / PO-xxx"/></F>
            <F label="Alasan Retur *"><input value={form.reason} onChange={e=>setForm(f=>({...f,reason:e.target.value}))} style={inp} placeholder="Barang cacat, salah kirim..."/></F>
            <div style={{fontWeight:600,fontSize:12,color:C.muted,marginBottom:-4}}>Item yang Diretur</div>
            {form.items.map((item,i)=>(
              <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 60px 100px 24px",gap:6,alignItems:"end"}}>
                <F label={i===0?"Produk":""}><select value={item.productId} onChange={e=>updateItem(i,"productId",e.target.value)} style={sel}><option value="">—</option>{products.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></F>
                <F label={i===0?"Qty":""}><input type="number" min="1" value={item.qty} onChange={e=>updateItem(i,"qty",e.target.value)} style={{...inp,textAlign:"center"}}/></F>
                <F label={i===0?"Harga":""}><input type="number" value={item.price} onChange={e=>updateItem(i,"price",e.target.value)} style={inp}/></F>
                <button onClick={()=>setForm(f=>({...f,items:f.items.filter((_,idx)=>idx!==i)}))} style={{...btnSm(C.danger),padding:"4px 6px",marginBottom:0,alignSelf:"flex-end",height:36}}>✕</button>
              </div>
            ))}
            <button onClick={addItem} style={{...btn("#64748B"),padding:"6px 12px",fontSize:12}}>+ Tambah Item</button>
            <div style={{background:C.surface,borderRadius:8,padding:"9px 13px",display:"flex",justifyContent:"space-between"}}>
              <span style={{fontSize:13,color:C.muted}}>Total Retur</span>
              <span style={{fontWeight:700,color:C.danger}}>{fmt(form.items.reduce((s,it)=>s+it.qty*it.price,0))}</span>
            </div>
            <button onClick={submitRetur} style={{...btn(C.danger),padding:"10px"}}>↩️ Ajukan Retur</button>
          </div>
        </div>

        {/* Daftar Retur */}
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:15}}>Daftar Retur</div>
            <ExportBtn onExcel={()=>exportExcel([{name:"Retur",...buildReturnExport(returns)}])} onPdf={()=>exportPDF("Laporan Retur",buildReturnExport(returns).html)}/>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["No Retur","Tgl","Tipe","Pihak","Items","Total","Alasan","Status","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {returns.length===0?<tr><td colSpan={9} style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada data retur</td></tr>
                :returns.slice().reverse().map(r=>(
                <tr key={r.id} style={{borderBottom:`1px solid ${C.border}`}}>
                  <td style={tdS}><span style={{color:C.primary,fontWeight:700,fontSize:12}}>{r.id}</span></td>
                  <td style={tdS}>{r.date}</td>
                  <td style={tdS}><span style={{background:r.type==="penjualan"?C.primaryL:C.purpleL,color:r.type==="penjualan"?C.primary:C.purple,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:500}}>{r.type}</span></td>
                  <td style={tdS}><div style={{fontWeight:600,fontSize:12}}>{r.customer||r.supplier||"—"}</div></td>
                  <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{r.items.length} item</span></td>
                  <td style={tdS}><span style={{fontWeight:700,color:C.danger}}>{fmt(r.total)}</span></td>
                  <td style={tdS}><span style={{fontSize:12,color:C.muted,maxWidth:120,display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.reason}</span></td>
                  <td style={tdS}><Badge status={r.status}/></td>
                  <td style={tdS}>
                    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                      {r.status==="menunggu"&&<><button onClick={()=>setujuiRetur(r)} style={btnSm(C.success)}>✓</button><button onClick={()=>tolakRetur(r.id)} style={btnSm("#94A3B8")}>✕</button></>}
                      <button onClick={()=>setEditRetur({...r})} style={btnSm("#64748B")}>✏️</button>
                      <button onClick={()=>setConfirmDel(r.id)} style={btnSm(C.danger)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {editRetur&&<Modal onClose={()=>setEditRetur(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:16}}>✏️ Edit Retur</div>
        <div style={{display:"grid",gap:12}}>
          <F label="Status"><select value={editRetur.status} onChange={e=>setEditRetur(p=>({...p,status:e.target.value}))} style={sel}>{["menunggu","disetujui","ditolak"].map(s=><option key={s}>{s}</option>)}</select></F>
          <F label="Alasan"><input value={editRetur.reason} onChange={e=>setEditRetur(p=>({...p,reason:e.target.value}))} style={inp}/></F>
          <F label="Catatan"><input value={editRetur.notes||""} onChange={e=>setEditRetur(p=>({...p,notes:e.target.value}))} style={inp}/></F>
          <button onClick={()=>{setReturns(prev=>prev.map(r=>r.id===editRetur.id?editRetur:r));setEditRetur(null);}} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg="Hapus data retur ini?" onConfirm={()=>deleteRetur(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

// ─── MULTI-GUDANG PAGE ────────────────────────────────────────────────────────
function GudangPage({warehouses,setWarehouses,products,setProducts,setMovements}){
  const [editWH,setEditWH]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [transferForm,setTransferForm]=useState({productId:"",qty:"",from:"",to:"",date:today(),note:""});
  const [transfers,setTransfers]=useState([]);
  const [tab,setTab]=useState("gudang");

  const saveWH=(w)=>{
    if(w._new){setWarehouses(prev=>[...prev,{...w,id:Date.now(),used:0,_new:undefined}]);}
    else{setWarehouses(prev=>prev.map(x=>x.id===w.id?w:x));}
    setEditWH(null);
  };
  const deleteWH=(id)=>{setWarehouses(prev=>prev.filter(w=>w.id!==id));setConfirmDel(null);};

  const doTransfer=()=>{
    const prod=products.find(p=>p.id===Number(transferForm.productId));
    if(!prod||!transferForm.qty||!transferForm.from||!transferForm.to)return;
    if(transferForm.from===transferForm.to){alert("Gudang asal dan tujuan harus berbeda!");return;}
    const qty=Number(transferForm.qty);
    if(qty>prod.stock){alert(`Stok tidak cukup! Stok saat ini: ${prod.stock}`);return;}
    const before=prod.stock;
    const ref=`TRF-${Date.now()}`;
    // stok produk tidak berubah (hanya lokasi), catat mutasi transfer
    setMovements(prev=>[
      ...prev,
      {id:Date.now(),date:transferForm.date,type:"out",subtype:"transfer",product:prod.name,productId:prod.id,qty,ref,note:`Transfer dari ${transferForm.from} ke ${transferForm.to}`,before,after:before},
      {id:Date.now()+1,date:transferForm.date,type:"in",subtype:"transfer",product:prod.name,productId:prod.id,qty,ref,note:`Terima transfer dari ${transferForm.from} ke ${transferForm.to}`,before,after:before},
    ]);
    setTransfers(prev=>[...prev,{id:ref,date:transferForm.date,product:prod.name,qty,from:transferForm.from,to:transferForm.to,note:transferForm.note}]);
    setTransferForm({productId:"",qty:"",from:"",to:"",date:today(),note:""});
    alert(`✅ Transfer ${qty} unit ${prod.name} dari ${transferForm.from} ke ${transferForm.to} berhasil!`);
  };

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};
  const warehouseNames=warehouses.map(w=>w.name);

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:20}}>
        {warehouses.map(w=>{
          const pct=w.capacity>0?Math.round((w.used/w.capacity)*100):0;
          return(
            <div key={w.id} style={{...card,borderTop:`4px solid ${pct>80?C.danger:pct>50?C.warning:C.success}`}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                <div><div style={{fontWeight:700,fontSize:15}}>{w.name}</div><div style={{fontSize:12,color:C.muted}}>{w.location}</div></div>
                <div style={{textAlign:"right"}}><div style={{fontWeight:700,fontSize:16,color:pct>80?C.danger:C.success}}>{pct}%</div><div style={{fontSize:11,color:C.muted}}>terisi</div></div>
              </div>
              <div style={{background:C.surface,borderRadius:6,height:8,overflow:"hidden",marginBottom:8}}>
                <div style={{width:`${pct}%`,height:"100%",background:pct>80?C.danger:pct>50?C.warning:C.success,borderRadius:6}}/>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:C.muted}}>
                <span>PIC: {w.pic||"—"}</span><span>{w.used}/{w.capacity} unit</span>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{display:"flex",gap:4,marginBottom:16,background:"#FFF",borderRadius:12,padding:6,border:`1px solid ${C.border}`,width:"fit-content",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        {[{id:"gudang",l:"🏭 Data Gudang"},{id:"transfer",l:"🔄 Transfer Stok"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?C.primary:"transparent",color:tab===t.id?"#FFF":C.muted}}>{t.l}</button>
        ))}
      </div>

      {tab==="gudang"&&(
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:15}}>Data Gudang</div>
            <button onClick={()=>setEditWH({_new:true,name:"",location:"",pic:"",capacity:0,used:0})} style={btn(C.primary)}>+ Tambah Gudang</button>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["Nama Gudang","Lokasi","PIC","Kapasitas","Terisi","Tersisa","Utilisasi","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {warehouses.map(w=>{
                const sisa=w.capacity-w.used;
                const pct=w.capacity>0?Math.round((w.used/w.capacity)*100):0;
                return(
                  <tr key={w.id} style={{borderBottom:`1px solid ${C.border}`}}>
                    <td style={tdS}><div style={{fontWeight:700}}>{w.name}</div></td>
                    <td style={tdS}><div style={{fontSize:12,color:C.muted}}>{w.location}</div></td>
                    <td style={tdS}>{w.pic||"—"}</td>
                    <td style={tdS}>{w.capacity} unit</td>
                    <td style={tdS}><span style={{fontWeight:600,color:pct>80?C.danger:C.text}}>{w.used}</span></td>
                    <td style={tdS}><span style={{color:sisa<10?C.danger:C.success,fontWeight:600}}>{sisa}</span></td>
                    <td style={tdS}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{flex:1,background:C.surface,borderRadius:4,height:8,minWidth:80}}><div style={{width:`${pct}%`,height:"100%",background:pct>80?C.danger:pct>50?C.warning:C.success,borderRadius:4}}/></div>
                        <span style={{fontSize:11,fontWeight:600,color:pct>80?C.danger:C.muted}}>{pct}%</span>
                      </div>
                    </td>
                    <td style={tdS}><div style={{display:"flex",gap:5}}><button onClick={()=>setEditWH({...w})} style={btnSm("#64748B")}>✏️</button><button onClick={()=>setConfirmDel(w.id)} style={btnSm(C.danger)}>🗑️</button></div></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {tab==="transfer"&&(
        <div style={{display:"grid",gridTemplateColumns:"360px 1fr",gap:16,alignItems:"start"}}>
          <div style={{...card,borderTop:`4px solid ${C.primary}`}}>
            <div style={{fontWeight:700,fontSize:15,color:C.primary,marginBottom:4}}>🔄 Transfer Antar Gudang</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Pindahkan stok dari satu gudang ke gudang lain</div>
            <div style={{display:"grid",gap:12}}>
              <F label="Produk *"><select value={transferForm.productId} onChange={e=>setTransferForm(f=>({...f,productId:e.target.value}))} style={sel}><option value="">— Pilih produk —</option>{products.map(p=><option key={p.id} value={p.id}>{p.name} (Stok: {p.stock})</option>)}</select></F>
              <F label="Jumlah *"><input type="number" min="1" value={transferForm.qty} onChange={e=>setTransferForm(f=>({...f,qty:e.target.value}))} style={inp} placeholder="0"/></F>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <F label="Dari Gudang *"><select value={transferForm.from} onChange={e=>setTransferForm(f=>({...f,from:e.target.value}))} style={sel}><option value="">— Asal —</option>{warehouseNames.map(w=><option key={w}>{w}</option>)}</select></F>
                <F label="Ke Gudang *"><select value={transferForm.to} onChange={e=>setTransferForm(f=>({...f,to:e.target.value}))} style={sel}><option value="">— Tujuan —</option>{warehouseNames.map(w=><option key={w}>{w}</option>)}</select></F>
              </div>
              <F label="Tanggal"><input type="date" value={transferForm.date} onChange={e=>setTransferForm(f=>({...f,date:e.target.value}))} style={inp}/></F>
              <F label="Catatan"><input value={transferForm.note} onChange={e=>setTransferForm(f=>({...f,note:e.target.value}))} style={inp} placeholder="Keterangan..."/></F>
              <button onClick={doTransfer} style={{...btn(C.primary),padding:"11px"}}>🔄 Proses Transfer</button>
            </div>
          </div>
          <div style={card}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:14}}>Riwayat Transfer</div>
            {transfers.length===0?<div style={{textAlign:"center",padding:32,color:C.muted}}>Belum ada riwayat transfer</div>
              :<table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:C.surface}}>{["Tgl","Produk","Qty","Dari","Ke","Ref","Catatan"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
                <tbody>{transfers.slice().reverse().map(t=>(
                  <tr key={t.id} style={{borderBottom:`1px solid ${C.border}`}}>
                    <td style={tdS}>{t.date}</td><td style={tdS}><b>{t.product}</b></td><td style={tdS}><b style={{color:C.primary}}>{t.qty}</b></td>
                    <td style={tdS}><span style={{background:C.dangerL,color:C.danger,padding:"2px 8px",borderRadius:10,fontSize:11,fontWeight:500}}>{t.from}</span></td>
                    <td style={tdS}><span style={{background:C.successL,color:C.success,padding:"2px 8px",borderRadius:10,fontSize:11,fontWeight:500}}>{t.to}</span></td>
                    <td style={tdS}><span style={{color:C.primary,fontWeight:600,fontSize:11}}>{t.id}</span></td>
                    <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{t.note}</span></td>
                  </tr>
                ))}</tbody>
              </table>
            }
          </div>
        </div>
      )}

      {editWH&&<Modal onClose={()=>setEditWH(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:16}}>{editWH._new?"➕ Tambah Gudang":"✏️ Edit Gudang"}</div>
        <div style={{display:"grid",gap:12}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Nama Gudang *"><input value={editWH.name} onChange={e=>setEditWH(p=>({...p,name:e.target.value}))} style={inp} placeholder="Gudang A..."/></F>
            <F label="PIC"><input value={editWH.pic||""} onChange={e=>setEditWH(p=>({...p,pic:e.target.value}))} style={inp} placeholder="Nama penanggung jawab"/></F>
          </div>
          <F label="Lokasi"><input value={editWH.location} onChange={e=>setEditWH(p=>({...p,location:e.target.value}))} style={inp} placeholder="Lantai 1, Gedung..."/></F>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Kapasitas (unit)"><input type="number" min="0" value={editWH.capacity} onChange={e=>setEditWH(p=>({...p,capacity:Number(e.target.value)}))} style={inp}/></F>
            <F label="Terisi (unit)"><input type="number" min="0" value={editWH.used} onChange={e=>setEditWH(p=>({...p,used:Number(e.target.value)}))} style={inp}/></F>
          </div>
          <button onClick={()=>saveWH(editWH)} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg={`Hapus gudang "${warehouses.find(w=>w.id===confirmDel)?.name}"?`} onConfirm={()=>deleteWH(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

// ─── DISKON PAGE ──────────────────────────────────────────────────────────────
function DiskonPage({discounts,setDiscounts,customers,products}){
  const [editDisc,setEditDisc]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);

  const saveDisc=(d)=>{
    if(d._new){setDiscounts(prev=>[...prev,{...d,id:Date.now(),_new:undefined}]);}
    else{setDiscounts(prev=>prev.map(x=>x.id===d.id?d:x));}
    setEditDisc(null);
  };
  const deleteDisc=(id)=>{setDiscounts(prev=>prev.filter(d=>d.id!==id));setConfirmDel(null);};
  const toggleActive=(id)=>setDiscounts(prev=>prev.map(d=>d.id===id?{...d,active:!d.active}:d));

  // Simulator
  const [simCustomer,setSimCustomer]=useState("");
  const [simProduct,setSimProduct]=useState("");
  const [simAmount,setSimAmount]=useState("");
  const simResult=useMemo(()=>{
    if(!simAmount)return null;
    const amount=Number(simAmount);
    const custId=simCustomer?Number(simCustomer):null;
    const prodId=simProduct?Number(simProduct):null;
    const applicable=discounts.filter(d=>{
      if(!d.active)return false;
      if(d.minOrder>0&&amount<d.minOrder)return false;
      if(d.customerId&&d.customerId!==custId)return false;
      if(d.productId&&d.productId!==prodId)return false;
      return true;
    });
    const totalDisc=applicable.reduce((s,d)=>s+(d.type==="persen"?amount*d.value/100:d.value),0);
    return{applicable,totalDisc,finalAmount:Math.max(0,amount-totalDisc)};
  },[discounts,simCustomer,simProduct,simAmount]);

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total Diskon",v:discounts.length,c:C.primary,i:"🏷️"},
          {l:"Aktif",v:discounts.filter(d=>d.active).length,c:C.success,i:"✅"},
          {l:"Diskon Persen",v:discounts.filter(d=>d.type==="persen").length,c:C.warning,i:"%"},
          {l:"Diskon Nominal",v:discounts.filter(d=>d.type==="nominal").length,c:C.purple,i:"Rp"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:18,fontWeight:700,color:k.c}}>{k.v}</div></div><span style={{fontSize:24}}>{k.i}</span></div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 320px",gap:16,alignItems:"start"}}>
        {/* Daftar Diskon */}
        <div style={card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:15}}>Daftar Diskon & Harga Khusus</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setEditDisc({_new:true,name:"",type:"persen",value:0,minOrder:0,customerId:null,productId:null,active:true,desc:""})} style={btn(C.primary)}>+ Tambah Diskon</button>
              <ExportBtn onExcel={()=>exportExcel([{name:"Diskon",...buildDiscountExport(discounts,customers,products)}])} onPdf={()=>exportPDF("Daftar Diskon",buildDiscountExport(discounts,customers,products).html)}/>
            </div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["Nama Diskon","Tipe","Nilai","Min Order","Berlaku Untuk","Status","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {discounts.map(d=>{
                const custName=d.customerId?customers.find(c=>c.id===d.customerId)?.name:"Semua pelanggan";
                const prodName=d.productId?products.find(p=>p.id===d.productId)?.name:"Semua produk";
                return(
                  <tr key={d.id} style={{borderBottom:`1px solid ${C.border}`,opacity:d.active?1:0.5}}>
                    <td style={tdS}><div style={{fontWeight:700}}>{d.name}</div><div style={{fontSize:11,color:C.muted}}>{d.desc}</div></td>
                    <td style={tdS}><span style={{background:d.type==="persen"?C.warningL:C.purpleL,color:d.type==="persen"?"#92400E":C.purple,padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:600}}>{d.type==="persen"?"Persen":"Nominal"}</span></td>
                    <td style={tdS}><span style={{fontWeight:700,fontSize:15,color:C.success}}>{d.type==="persen"?`${d.value}%`:fmt(d.value)}</span></td>
                    <td style={tdS}>{d.minOrder>0?fmt(d.minOrder):<span style={{color:C.muted,fontSize:12}}>—</span>}</td>
                    <td style={tdS}><div style={{fontSize:12}}><div style={{color:C.primary}}>{custName}</div><div style={{color:C.muted}}>{prodName}</div></div></td>
                    <td style={tdS}><Badge status={d.active?"aktif":"nonaktif"}/></td>
                    <td style={tdS}>
                      <div style={{display:"flex",gap:5}}>
                        <button onClick={()=>toggleActive(d.id)} style={btnSm(d.active?"#94A3B8":C.success)}>{d.active?"Nonaktif":"Aktifkan"}</button>
                        <button onClick={()=>setEditDisc({...d})} style={btnSm("#64748B")}>✏️</button>
                        <button onClick={()=>setConfirmDel(d.id)} style={btnSm(C.danger)}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Simulator Diskon */}
        <div style={{...card,borderTop:`4px solid ${C.success}`}}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:4,color:C.success}}>🧮 Simulator Diskon</div>
          <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Cek diskon yang berlaku untuk transaksi</div>
          <div style={{display:"grid",gap:12}}>
            <F label="Total Order (IDR)"><input type="number" min="0" value={simAmount} onChange={e=>setSimAmount(e.target.value)} style={inp} placeholder="0"/></F>
            <F label="Pelanggan"><select value={simCustomer} onChange={e=>setSimCustomer(e.target.value)} style={sel}><option value="">Semua / Umum</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></F>
            <F label="Produk"><select value={simProduct} onChange={e=>setSimProduct(e.target.value)} style={sel}><option value="">Semua produk</option>{products.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></F>
            {simResult&&(
              <div style={{background:C.surface,borderRadius:10,padding:14}}>
                <div style={{fontWeight:600,fontSize:13,marginBottom:10}}>Diskon yang Berlaku ({simResult.applicable.length})</div>
                {simResult.applicable.length===0?<div style={{color:C.muted,fontSize:13}}>Tidak ada diskon berlaku</div>
                  :simResult.applicable.map(d=>(
                    <div key={d.id} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${C.border}`,fontSize:13}}>
                      <span>{d.name}</span>
                      <span style={{color:C.success,fontWeight:600}}>-{d.type==="persen"?`${d.value}%`:fmt(d.value)}</span>
                    </div>
                  ))
                }
                <div style={{marginTop:10,paddingTop:10,borderTop:`2px solid ${C.border}`}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:4}}><span style={{color:C.muted}}>Total Diskon</span><span style={{fontWeight:700,color:C.danger}}>-{fmt(simResult.totalDisc)}</span></div>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:14}}><span style={{fontWeight:700}}>Total Bayar</span><span style={{fontWeight:800,color:C.primary,fontSize:16}}>{fmt(simResult.finalAmount)}</span></div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {editDisc&&<Modal onClose={()=>setEditDisc(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editDisc._new?"➕ Tambah Diskon":"✏️ Edit Diskon"}</div>
        <div style={{display:"grid",gap:13}}>
          <F label="Nama Diskon *"><input value={editDisc.name} onChange={e=>setEditDisc(p=>({...p,name:e.target.value}))} style={inp} placeholder="Nama diskon..."/></F>
          <F label="Deskripsi"><input value={editDisc.desc||""} onChange={e=>setEditDisc(p=>({...p,desc:e.target.value}))} style={inp} placeholder="Keterangan diskon..."/></F>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Tipe Diskon"><select value={editDisc.type} onChange={e=>setEditDisc(p=>({...p,type:e.target.value}))} style={sel}><option value="persen">Persen (%)</option><option value="nominal">Nominal (Rp)</option></select></F>
            <F label={editDisc.type==="persen"?"Nilai (%)":"Nilai (IDR)"}><input type="number" min="0" value={editDisc.value} onChange={e=>setEditDisc(p=>({...p,value:Number(e.target.value)}))} style={inp}/></F>
          </div>
          <F label="Minimum Order (IDR, 0 = tidak ada)"><input type="number" min="0" value={editDisc.minOrder} onChange={e=>setEditDisc(p=>({...p,minOrder:Number(e.target.value)}))} style={inp}/></F>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Khusus Pelanggan"><select value={editDisc.customerId||""} onChange={e=>setEditDisc(p=>({...p,customerId:e.target.value?Number(e.target.value):null}))} style={sel}><option value="">Semua pelanggan</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></F>
            <F label="Khusus Produk"><select value={editDisc.productId||""} onChange={e=>setEditDisc(p=>({...p,productId:e.target.value?Number(e.target.value):null}))} style={sel}><option value="">Semua produk</option>{products.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></F>
          </div>
          <F label="Status"><select value={editDisc.active?"aktif":"nonaktif"} onChange={e=>setEditDisc(p=>({...p,active:e.target.value==="aktif"}))} style={sel}><option value="aktif">Aktif</option><option value="nonaktif">Non-Aktif</option></select></F>
          <button onClick={()=>saveDisc(editDisc)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg="Hapus diskon ini?" onConfirm={()=>deleteDisc(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

// ─── ABSENSI PAGE ─────────────────────────────────────────────────────────────
function AbsensiPage({attendance,setAttendance,employees,payrolls,setPayrolls}){
  const [tab,setTab]=useState("absensi");
  const [filterDate,setFilterDate]=useState(today());
  const [filterEmp,setFilterEmp]=useState("");
  const [editRec,setEditRec]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [form,setForm]=useState({employeeId:"",date:today(),checkIn:"08:00",checkOut:"17:00",status:"hadir",note:""});

  const addAbsensi=()=>{
    if(!form.employeeId)return;
    const exists=attendance.find(a=>a.employeeId===Number(form.employeeId)&&a.date===form.date);
    if(exists){alert("Absensi karyawan ini pada tanggal tersebut sudah ada!");return;}
    setAttendance(prev=>[...prev,{...form,id:uid(),employeeId:Number(form.employeeId)}]);
    setForm(f=>({...f,employeeId:"",note:""}));
  };
  const saveRec=(r)=>{setAttendance(prev=>prev.map(a=>a.id===r.id?r:a));setEditRec(null);};
  const deleteRec=(id)=>{setAttendance(prev=>prev.filter(a=>a.id!==id));setConfirmDel(null);};

  const filtered=attendance.filter(a=>{
    if(filterDate&&a.date!==filterDate)return false;
    if(filterEmp&&a.employeeId!==Number(filterEmp))return false;
    return true;
  });

  // Rekap per karyawan
  const rekapEmp=useMemo(()=>employees.map(e=>{
    const recs=attendance.filter(a=>a.employeeId===e.id);
    return{...e,hadir:recs.filter(a=>a.status==="hadir").length,terlambat:recs.filter(a=>a.status==="terlambat").length,izin:recs.filter(a=>a.status==="izin").length,sakit:recs.filter(a=>a.status==="sakit").length,alpha:recs.filter(a=>a.status==="alpha").length,total:recs.length};
  }),[attendance,employees]);

  // Sinkron ke payroll
  const syncPayroll=(empId,month)=>{
    const recs=attendance.filter(a=>a.employeeId===empId&&a.date.startsWith(month));
    const alpha=recs.filter(a=>a.status==="alpha").length;
    const terlambat=recs.filter(a=>a.status==="terlambat").length;
    const pay=payrolls.find(p=>p.employeeId===empId&&p.month===month);
    if(!pay){alert("Proses gaji bulan ini dulu di modul HR & Gaji!");return;}
    const emp=employees.find(e=>e.id===empId);
    if(!emp)return;
    const potonganAlpha=alpha*(emp.gajiPokok/25);
    const potonganTerlambat=terlambat*50000;
    const potonganBaru=potonganAlpha+potonganTerlambat;
    setPayrolls(prev=>prev.map(p=>{
      if(p.id!==pay.id)return p;
      const totalPot=p.potonganBPJS+p.potonganPPh+potonganBaru;
      return{...p,potonganLain:potonganBaru,potonganLainKet:`Alpha ${alpha}hr, Terlambat ${terlambat}hr`,totalPotongan:totalPot,takehome:p.gajiPokok+p.totalTunjangan-totalPot};
    }));
    alert(`✅ Sinkron berhasil!\nPotongan alpha (${alpha} hari): ${fmt(potonganAlpha)}\nPotongan terlambat (${terlambat}x): ${fmt(potonganTerlambat)}`);
  };

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Hadir Hari Ini",v:attendance.filter(a=>a.date===today()&&a.status==="hadir").length,c:C.success,i:"✅"},
          {l:"Terlambat",v:attendance.filter(a=>a.date===today()&&a.status==="terlambat").length,c:C.warning,i:"⏰"},
          {l:"Izin",v:attendance.filter(a=>a.date===today()&&a.status==="izin").length,c:C.primary,i:"📋"},
          {l:"Sakit",v:attendance.filter(a=>a.date===today()&&a.status==="sakit").length,c:"#7C3AED",i:"🏥"},
          {l:"Alpha",v:attendance.filter(a=>a.date===today()&&a.status==="alpha").length,c:C.danger,i:"❌"},
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`4px solid ${k.c}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:4}}>{k.l}</div><div style={{fontSize:22,fontWeight:700,color:k.c}}>{k.v}</div></div><span style={{fontSize:22}}>{k.i}</span></div>
          </div>
        ))}
      </div>

      <div style={{display:"flex",gap:4,marginBottom:16,background:"#FFF",borderRadius:12,padding:6,border:`1px solid ${C.border}`,width:"fit-content",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        {[{id:"absensi",l:"📅 Input Absensi"},{id:"rekap",l:"📊 Rekap Karyawan"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?C.primary:"transparent",color:tab===t.id?"#FFF":C.muted}}>{t.l}</button>
        ))}
      </div>

      {tab==="absensi"&&(
        <div style={{display:"grid",gridTemplateColumns:"340px 1fr",gap:16,alignItems:"start"}}>
          <div style={{...card,borderTop:`4px solid ${C.primary}`}}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:16,color:C.primary}}>📅 Input Absensi</div>
            <div style={{display:"grid",gap:12}}>
              <F label="Karyawan *"><select value={form.employeeId} onChange={e=>setForm(f=>({...f,employeeId:e.target.value}))} style={sel}><option value="">— Pilih —</option>{employees.filter(e=>e.status==="aktif").map(e=><option key={e.id} value={e.id}>{e.name}</option>)}</select></F>
              <F label="Tanggal"><input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))} style={inp}/></F>
              <F label="Status"><select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={sel}>{["hadir","terlambat","izin","sakit","alpha"].map(s=><option key={s}>{s}</option>)}</select></F>
              {(form.status==="hadir"||form.status==="terlambat")&&(
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <F label="Check In"><input type="time" value={form.checkIn} onChange={e=>setForm(f=>({...f,checkIn:e.target.value}))} style={inp}/></F>
                  <F label="Check Out"><input type="time" value={form.checkOut} onChange={e=>setForm(f=>({...f,checkOut:e.target.value}))} style={inp}/></F>
                </div>
              )}
              <F label="Catatan"><input value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} style={inp} placeholder="Keterangan (opsional)"/></F>
              <button onClick={addAbsensi} style={{...btn(C.primary),padding:"10px"}}>💾 Simpan Absensi</button>
            </div>
          </div>

          <div style={card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontWeight:700,fontSize:15}}>Riwayat Absensi</div>
              <div style={{display:"flex",gap:8}}>
                <input type="date" value={filterDate} onChange={e=>setFilterDate(e.target.value)} style={{...inp,width:150}}/>
                <select value={filterEmp} onChange={e=>setFilterEmp(e.target.value)} style={{...sel,width:180}}><option value="">Semua karyawan</option>{employees.map(e=><option key={e.id} value={e.id}>{e.name}</option>)}</select>
                <button onClick={()=>{setFilterDate("");setFilterEmp("");}} style={{...btn("#94A3B8"),padding:"7px 12px",fontSize:12}}>Reset</button>
              </div>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr style={{background:C.surface}}>{["Tanggal","Karyawan","Status","Check In","Check Out","Jam Kerja","Catatan","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
              <tbody>
                {filtered.length===0?<tr><td colSpan={8} style={{textAlign:"center",padding:28,color:C.muted}}>Tidak ada data</td></tr>
                  :filtered.slice().sort((a,b)=>b.date.localeCompare(a.date)).map(a=>{
                    const emp=employees.find(e=>e.id===a.employeeId);
                    const jamKerja=a.checkIn&&a.checkOut?`${Math.floor((new Date(`2000-01-01T${a.checkOut}`)-new Date(`2000-01-01T${a.checkIn}`))/(1000*60*60))} jam`:"—";
                    return(
                      <tr key={a.id} style={{borderBottom:`1px solid ${C.border}`}}>
                        <td style={tdS}>{a.date}</td>
                        <td style={tdS}><div style={{fontWeight:600}}>{emp?.name||"—"}</div><div style={{fontSize:11,color:C.muted}}>{emp?.dept}</div></td>
                        <td style={tdS}><Badge status={a.status}/></td>
                        <td style={tdS}><span style={{fontWeight:500,color:C.success}}>{a.checkIn||"—"}</span></td>
                        <td style={tdS}><span style={{fontWeight:500,color:C.muted}}>{a.checkOut||"—"}</span></td>
                        <td style={tdS}><span style={{fontWeight:600}}>{jamKerja}</span></td>
                        <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{a.note||"—"}</span></td>
                        <td style={tdS}><div style={{display:"flex",gap:5}}><button onClick={()=>setEditRec({...a})} style={btnSm("#64748B")}>✏️</button><button onClick={()=>setConfirmDel(a.id)} style={btnSm(C.danger)}>🗑️</button></div></td>
                      </tr>
                    );
                  })
                }
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab==="rekap"&&(
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>Rekap Absensi Karyawan</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:C.surface}}>{["Karyawan","Dept","Total Hari","Hadir","Terlambat","Izin","Sakit","Alpha","% Kehadiran","Sinkron Gaji"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
            <tbody>
              {rekapEmp.map(e=>{
                const pct=e.total>0?Math.round(((e.hadir+e.terlambat)/e.total)*100):0;
                return(
                  <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`}}>
                    <td style={tdS}><div style={{fontWeight:700}}>{e.name}</div></td>
                    <td style={tdS}><span style={{background:C.primaryL,color:C.primary,padding:"2px 8px",borderRadius:12,fontSize:11}}>{e.dept}</span></td>
                    <td style={tdS}><b>{e.total}</b></td>
                    <td style={tdS}><span style={{color:C.success,fontWeight:600}}>{e.hadir}</span></td>
                    <td style={tdS}><span style={{color:C.warning,fontWeight:600}}>{e.terlambat}</span></td>
                    <td style={tdS}><span style={{color:C.primary,fontWeight:600}}>{e.izin}</span></td>
                    <td style={tdS}><span style={{color:"#7C3AED",fontWeight:600}}>{e.sakit}</span></td>
                    <td style={tdS}><span style={{color:C.danger,fontWeight:600}}>{e.alpha}</span></td>
                    <td style={tdS}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{flex:1,background:C.surface,borderRadius:4,height:8,minWidth:60}}><div style={{width:`${pct}%`,height:"100%",background:pct<75?C.danger:pct<90?C.warning:C.success,borderRadius:4}}/></div>
                        <span style={{fontSize:12,fontWeight:600,color:pct<75?C.danger:pct<90?C.warning:C.success}}>{pct}%</span>
                      </div>
                    </td>
                    <td style={tdS}><button onClick={()=>syncPayroll(e.id,today().slice(0,7))} style={btnSm(C.warning)}>🔗 Sinkron</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editRec&&<Modal onClose={()=>setEditRec(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:16}}>✏️ Edit Absensi</div>
        <div style={{display:"grid",gap:12}}>
          <F label="Status"><select value={editRec.status} onChange={e=>setEditRec(p=>({...p,status:e.target.value}))} style={sel}>{["hadir","terlambat","izin","sakit","alpha"].map(s=><option key={s}>{s}</option>)}</select></F>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <F label="Check In"><input type="time" value={editRec.checkIn||""} onChange={e=>setEditRec(p=>({...p,checkIn:e.target.value}))} style={inp}/></F>
            <F label="Check Out"><input type="time" value={editRec.checkOut||""} onChange={e=>setEditRec(p=>({...p,checkOut:e.target.value}))} style={inp}/></F>
          </div>
          <F label="Catatan"><input value={editRec.note||""} onChange={e=>setEditRec(p=>({...p,note:e.target.value}))} style={inp}/></F>
          <button onClick={()=>saveRec(editRec)} style={{...btn(C.primary),padding:"11px"}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg="Hapus catatan absensi ini?" onConfirm={()=>deleteRec(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

// ─── NOTIFIKASI COMPONENT (shown in all pages) ────────────────────────────────
function NotifPanel({products,invoices,pos,onClose}){
  const alerts=useMemo(()=>{
    const list=[];
    // Stok kritis
    products.filter(p=>p.stock<=p.minStock).forEach(p=>list.push({id:`stok-${p.id}`,type:"danger",icon:"⚠️",title:"Stok Kritis",msg:`${p.name} tinggal ${p.stock} unit (min: ${p.minStock})`,page:"inventory"}));
    // Invoice jatuh tempo
    const todayD=new Date(today());
    invoices.filter(i=>i.status!=="paid").forEach(i=>{
      const due=new Date(i.dueDate);
      const diff=Math.ceil((due-todayD)/(1000*60*60*24));
      if(diff<=3)list.push({id:`inv-${i.id}`,type:diff<0?"danger":"warning",icon:"🧾",title:diff<0?"Invoice Lewat Jatuh Tempo":"Invoice Hampir Jatuh Tempo",msg:`${i.id} – ${i.customer} · ${fmt(i.grandTotal)} (${diff<0?`${Math.abs(diff)} hari lalu`:`${diff} hari lagi`})`,page:"invoices"});
    });
    // PO menunggu terima > 7 hari
    pos.filter(p=>p.status==="dipesan").forEach(p=>{
      const diff=Math.ceil((new Date(today())-new Date(p.date))/(1000*60*60*24));
      if(diff>7)list.push({id:`po-${p.id}`,type:"warning",icon:"📋",title:"PO Belum Diterima",msg:`${p.id} – ${p.supplierName} sudah ${diff} hari belum dikonfirmasi`,page:"po"});
    });
    return list;
  },[products,invoices,pos]);

  return(
    <div style={{position:"fixed",top:60,right:20,width:360,background:"#FFF",borderRadius:14,boxShadow:"0 12px 40px rgba(0,0,0,0.18)",border:`1px solid ${C.border}`,zIndex:500,overflow:"hidden"}}>
      <div style={{padding:"14px 16px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",background:C.surface}}>
        <div style={{fontWeight:700,fontSize:14}}>🔔 Notifikasi <span style={{background:C.danger,color:"#FFF",borderRadius:10,fontSize:11,padding:"1px 6px",fontWeight:700,marginLeft:6}}>{alerts.length}</span></div>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:16,color:C.muted}}>✕</button>
      </div>
      <div style={{maxHeight:400,overflowY:"auto"}}>
        {alerts.length===0
          ?<div style={{padding:32,textAlign:"center",color:C.muted}}><div style={{fontSize:32,marginBottom:8}}>✅</div>Tidak ada notifikasi</div>
          :alerts.map(a=>(
            <div key={a.id} style={{padding:"12px 16px",borderBottom:`1px solid ${C.border}`,background:a.type==="danger"?"#FFF9F9":a.type==="warning"?"#FFFDF0":"#F0FFF4"}}>
              <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
                <span style={{fontSize:18,flexShrink:0}}>{a.icon}</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:12,color:a.type==="danger"?C.danger:a.type==="warning"?"#92400E":C.success,marginBottom:3}}>{a.title}</div>
                  <div style={{fontSize:12,color:C.text}}>{a.msg}</div>
                </div>
              </div>
            </div>
          ))
        }
      </div>
    </div>
  );
}

// ─── DASHBOARD CHART PAGE ─────────────────────────────────────────────────────
function DashboardPage({salesData,products,orders,invoices,cashFlowEntries,employees,payrolls}){
  const maxRevenue=Math.max(...salesData.map(d=>d.revenue),1);
  const maxExp=Math.max(...salesData.map(d=>d.expenses),1);
  const maxBoth=Math.max(maxRevenue,maxExp);
  const totalRevenue=salesData.reduce((s,d)=>s+d.revenue,0);
  const totalExp=salesData.reduce((s,d)=>s+d.expenses,0);
  const totalOrders=salesData.reduce((s,d)=>s+d.orders,0);
  const avgMonthly=Math.round(totalRevenue/salesData.length);
  const lastMonth=salesData[salesData.length-1];
  const prevMonth=salesData[salesData.length-2];
  const growth=prevMonth?Math.round(((lastMonth.revenue-prevMonth.revenue)/prevMonth.revenue)*100):0;

  // Top produk
  const topProds=[...products].sort((a,b)=>(b.stock*b.price)-(a.stock*a.price)).slice(0,4);
  // Order status pie
  const statusMap=orders.reduce((a,o)=>{a[o.status]=(a[o.status]||0)+1;return a;},{});
  const statusColors={pending:C.warning,processing:C.primary,delivered:C.success,cancelled:C.danger};

  return(
    <div>
      {/* KPI Row */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
        {[
          {l:"Total Pendapatan YTD",v:fmt(totalRevenue),c:C.success,i:"💰",sub:`Rata-rata ${fmt(avgMonthly)}/bln`},
          {l:"Pertumbuhan Bulan Ini",v:`${growth>=0?"+":""}${growth}%`,c:growth>=0?C.success:C.danger,i:growth>=0?"📈":"📉",sub:`vs bulan lalu`},
          {l:"Total Pesanan",v:totalOrders,c:C.primary,i:"🛒",sub:`${salesData.length} bulan data`},
          {l:"Karyawan Aktif",v:employees.filter(e=>e.status==="aktif").length,c:"#7C3AED",i:"👥",sub:`Total gaji: ${fmt(payrolls.filter(p=>p.month===today().slice(0,7)).reduce((s,p)=>s+p.takehome,0))}`},
        ].map((k,i)=>(
          <div key={i} style={card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div><div style={{fontSize:11,color:C.muted,fontWeight:500,marginBottom:6}}>{k.l}</div><div style={{fontSize:22,fontWeight:700,color:k.c}}>{k.v}</div><div style={{fontSize:11,color:C.muted,marginTop:3}}>{k.sub}</div></div>
              <span style={{fontSize:28}}>{k.i}</span>
            </div>
          </div>
        ))}
      </div>

      {/* Main Chart */}
      <div style={{...card,marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div><div style={{fontWeight:700,fontSize:16}}>Grafik Penjualan & Pengeluaran</div><div style={{fontSize:12,color:C.muted}}>6 bulan terakhir</div></div>
          <div style={{display:"flex",gap:20,fontSize:12}}>
            <span style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:14,height:14,borderRadius:3,background:C.success,display:"inline-block"}}/>Pendapatan</span>
            <span style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:14,height:14,borderRadius:3,background:C.danger,display:"inline-block",opacity:0.75}}/>Pengeluaran</span>
          </div>
        </div>
        {/* Bar Chart */}
        <div style={{display:"flex",gap:6,alignItems:"flex-end",height:220,padding:"0 4px",borderBottom:`2px solid ${C.border}`}}>
          {salesData.map((d,i)=>{
            const revH=Math.round((d.revenue/maxBoth)*190);
            const expH=Math.round((d.expenses/maxBoth)*190);
            const laba=d.revenue-d.expenses;
            const fmtShort=v=>{if(v>=1000000000)return`${(v/1000000000).toFixed(1)}M`;if(v>=1000000)return`${(v/1000000).toFixed(0)}jt`;return`${(v/1000).toFixed(0)}rb`;};
            return(
              <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-end",height:"100%"}}>
                {/* Bars */}
                <div style={{display:"flex",gap:4,alignItems:"flex-end",width:"100%",justifyContent:"center"}}>
                  <div style={{position:"relative",display:"flex",flexDirection:"column",alignItems:"center"}}>
                    <div style={{fontSize:9,color:C.success,fontWeight:700,marginBottom:3,whiteSpace:"nowrap"}}>{fmtShort(d.revenue)}</div>
                    <div style={{width:22,height:revH,background:C.success,borderRadius:"4px 4px 0 0",minHeight:4,cursor:"pointer"}} title={`Pendapatan: ${fmt(d.revenue)}`}/>
                  </div>
                  <div style={{position:"relative",display:"flex",flexDirection:"column",alignItems:"center"}}>
                    <div style={{fontSize:9,color:C.danger,fontWeight:700,marginBottom:3,whiteSpace:"nowrap"}}>{fmtShort(d.expenses)}</div>
                    <div style={{width:22,height:expH,background:C.danger,borderRadius:"4px 4px 0 0",minHeight:4,opacity:0.8,cursor:"pointer"}} title={`Pengeluaran: ${fmt(d.expenses)}`}/>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
        {/* X-axis labels */}
        <div style={{display:"flex",gap:6,padding:"8px 4px 0"}}>
          {salesData.map((d,i)=>{
            const laba=d.revenue-d.expenses;
            return(
              <div key={i} style={{flex:1,textAlign:"center"}}>
                <div style={{fontSize:12,fontWeight:600,color:C.text}}>{MONTHS[parseInt(d.month.slice(5))-1]}</div>
                <div style={{fontSize:10,color:C.muted,marginTop:1}}>{d.orders} order</div>
                <div style={{fontSize:10,fontWeight:700,color:laba>=0?C.success:C.danger,marginTop:2}}>{laba>=0?"+":"-"}{Math.round(Math.abs(laba)/1000000)}jt</div>
              </div>
            );
          })}
        </div>
        {/* Summary footer */}
        <div style={{display:"flex",justifyContent:"space-between",marginTop:14,paddingTop:12,borderTop:`1px solid ${C.border}`,fontSize:12,color:C.muted}}>
          <span>Total Pendapatan: <b style={{color:C.success}}>{fmt(totalRevenue)}</b></span>
          <span>Total Pengeluaran: <b style={{color:C.danger}}>{fmt(totalExp)}</b></span>
          <span>Laba Bersih: <b style={{color:(totalRevenue-totalExp)>=0?C.success:C.danger}}>{fmt(totalRevenue-totalExp)}</b></span>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        {/* Order Status Chart */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>Status Pesanan</div>
          {Object.entries(statusMap).map(([st,cnt])=>{
            const pct=Math.round((cnt/orders.length)*100);
            return(
              <div key={st} style={{marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:5}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}><span style={{width:10,height:10,borderRadius:2,background:statusColors[st]||C.muted,display:"inline-block"}}/><Badge status={st}/></div>
                  <span style={{fontWeight:700}}>{cnt} <span style={{color:C.muted,fontWeight:400}}>({pct}%)</span></span>
                </div>
                <div style={{background:C.surface,borderRadius:5,height:8}}><div style={{width:`${pct}%`,height:"100%",background:statusColors[st]||C.muted,borderRadius:5}}/></div>
              </div>
            );
          })}
          <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",fontSize:13}}>
            <span style={{color:C.muted}}>Total Nilai Pesanan</span>
            <span style={{fontWeight:700,color:C.primary}}>{fmt(orders.reduce((s,o)=>s+o.total,0))}</span>
          </div>
        </div>

        {/* Top Produk */}
        <div style={card}>
          <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>Top Produk (Nilai Stok)</div>
          {topProds.map((p,i)=>{
            const maxVal=topProds[0].stock*topProds[0].price;
            const pct=Math.round(((p.stock*p.price)/maxVal)*100);
            return(
              <div key={p.id} style={{marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:5}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{width:20,height:20,borderRadius:5,background:C.primaryL,color:C.primary,fontWeight:700,fontSize:11,display:"flex",alignItems:"center",justifyContent:"center"}}>{i+1}</span>
                    <span style={{fontWeight:600}}>{p.name}</span>
                  </div>
                  <span style={{fontWeight:700,color:p.stock<=p.minStock?C.danger:C.text}}>{fmt(p.stock*p.price)}</span>
                </div>
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <div style={{flex:1,background:C.surface,borderRadius:4,height:7}}><div style={{width:`${pct}%`,height:"100%",background:p.stock<=p.minStock?C.danger:C.primary,borderRadius:4}}/></div>
                  <span style={{fontSize:11,color:p.stock<=p.minStock?C.danger:C.muted,fontWeight:500,minWidth:50,textAlign:"right"}}>{p.stock} unit</span>
                </div>
              </div>
            );
          })}
          <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",fontSize:13}}>
            <span style={{color:C.muted}}>Total Nilai Inventori</span>
            <span style={{fontWeight:700,color:C.primary}}>{fmt(products.reduce((s,p)=>s+p.stock*p.price,0))}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── USER & ROLE PAGE ─────────────────────────────────────────────────────────
function UserRolePage({users,setUsers,currentUser,setCurrentUser}){
  const [editUser,setEditUser]=useState(null);
  const [confirmDel,setConfirmDel]=useState(null);
  const [showPwd,setShowPwd]=useState(false);
  const [loginForm,setLoginForm]=useState({username:"",password:""});
  const [loginError,setLoginError]=useState("");

  const ROLES=[
    {id:"admin",label:"Admin",desc:"Akses penuh semua modul",color:"#1E3A5F",pages:["all"]},
    {id:"manager",label:"Manager",desc:"Semua kecuali user management",color:C.purple,pages:["dashboard","inventory","orders","invoices","cashflow","po","retur","gudang","diskon","hr","laporan","customers","absensi"]},
    {id:"kasir",label:"Kasir",desc:"Pesanan, invoice, diskon",color:C.primary,pages:["dashboard","orders","invoices","diskon","customers"]},
    {id:"gudang",label:"Gudang",desc:"Inventori, PO, gudang, retur",color:C.success,pages:["dashboard","inventory","po","retur","gudang"]},
  ];

  const saveUser=(u)=>{
    if(u._new){setUsers(prev=>[...prev,{...u,id:uid(),lastLogin:"-",_new:undefined,password:undefined}]);}
    else{setUsers(prev=>prev.map(x=>x.id===u.id?{...u,password:undefined}:x));}
    setEditUser(null);
  };
  const deleteUser=(id)=>{setUsers(prev=>prev.filter(u=>u.id!==id));setConfirmDel(null);};
  const toggleActive=(id)=>setUsers(prev=>prev.map(u=>u.id===id?{...u,active:!u.active}:u));

  const doLogin=()=>{
    const u=users.find(x=>x.username===loginForm.username&&x.active);
    if(!u){setLoginError("Username tidak ditemukan atau akun non-aktif");return;}
    // demo: password = username
    if(loginForm.password!==loginForm.username&&loginForm.password!=="gks123"){setLoginError("Password salah (hint: gunakan username atau 'gks123')");return;}
    setCurrentUser(u);
    setUsers(prev=>prev.map(x=>x.id===u.id?{...x,lastLogin:new Date().toLocaleString("id-ID")}:x));
    setLoginError("");
  };

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div>
      {/* Login Simulator */}
      <div style={{...card,marginBottom:16,borderTop:`4px solid ${C.primary}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>🔐 Simulasi Login</div>
            <div style={{fontSize:12,color:C.muted,marginBottom:14}}>Coba login sebagai user berbeda untuk melihat akses berbasis role</div>
            <div style={{display:"flex",gap:10,alignItems:"flex-end"}}>
              <div style={{flex:1}}><F label="Username"><input value={loginForm.username} onChange={e=>setLoginForm(f=>({...f,username:e.target.value}))} style={inp} placeholder="Coba: admin, kasir1, gudang1..."/></F></div>
              <div style={{flex:1}}><F label="Password"><input type={showPwd?"text":"password"} value={loginForm.password} onChange={e=>setLoginForm(f=>({...f,password:e.target.value}))} style={inp} placeholder="gks123 atau username"/></F></div>
              <button onClick={doLogin} style={{...btn(C.primary),height:38,marginBottom:0}}>Login</button>
              <button onClick={()=>setCurrentUser(null)} style={{...btn("#94A3B8"),height:38}}>Logout</button>
            </div>
            {loginError&&<div style={{color:C.danger,fontSize:12,marginTop:6}}>{loginError}</div>}
          </div>
          {currentUser&&(
            <div style={{marginLeft:20,background:C.surface,borderRadius:10,padding:"12px 16px",minWidth:200,border:`1px solid ${C.border}`}}>
              <div style={{fontSize:11,color:C.muted,marginBottom:4}}>Login sebagai</div>
              <div style={{fontWeight:700,fontSize:15}}>{currentUser.name}</div>
              <div style={{marginTop:4}}><Badge status={currentUser.role}/></div>
              <div style={{fontSize:11,color:C.muted,marginTop:6}}>{currentUser.email}</div>
              <div style={{fontSize:10,color:C.muted,marginTop:3}}>Login: {currentUser.lastLogin}</div>
            </div>
          )}
        </div>
      </div>

      {/* Role Cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:16}}>
        {ROLES.map(r=>(
          <div key={r.id} style={{...card,borderTop:`4px solid ${r.color}`}}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:4}}><Badge status={r.id}/></div>
            <div style={{fontSize:12,color:C.muted,marginBottom:8}}>{r.desc}</div>
            <div style={{fontSize:11,color:C.muted}}>{users.filter(u=>u.role===r.id).length} pengguna</div>
          </div>
        ))}
      </div>

      {/* User Table */}
      <div style={card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontWeight:700,fontSize:15}}>Manajemen User</div>
          <button onClick={()=>setEditUser({_new:true,username:"",name:"",role:"kasir",email:"",active:true})} style={btn(C.primary)}>+ Tambah User</button>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:C.surface}}>{["Username","Nama","Role","Email","Status","Login Terakhir","Akses Modul","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
          <tbody>
            {users.map(u=>{
              const role=ROLES.find(r=>r.id===u.role);
              return(
                <tr key={u.id} style={{borderBottom:`1px solid ${C.border}`,opacity:u.active?1:0.5}}>
                  <td style={tdS}><span style={{fontWeight:700,color:C.primary}}>{u.username}</span></td>
                  <td style={tdS}><div style={{fontWeight:600}}>{u.name}</div></td>
                  <td style={tdS}><Badge status={u.role}/></td>
                  <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{u.email}</span></td>
                  <td style={tdS}><Badge status={u.active?"aktif":"nonaktif"}/></td>
                  <td style={tdS}><span style={{fontSize:12,color:C.muted}}>{u.lastLogin}</span></td>
                  <td style={tdS}><span style={{fontSize:11,color:C.muted}}>{role?.pages[0]==="all"?"Semua modul":`${role?.pages.length} modul`}</span></td>
                  <td style={tdS}>
                    <div style={{display:"flex",gap:5}}>
                      <button onClick={()=>toggleActive(u.id)} style={btnSm(u.active?"#94A3B8":C.success)}>{u.active?"Nonaktif":"Aktifkan"}</button>
                      <button onClick={()=>setEditUser({...u})} style={btnSm("#64748B")}>✏️</button>
                      <button onClick={()=>setConfirmDel(u.id)} style={btnSm(C.danger)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {editUser&&<Modal onClose={()=>setEditUser(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editUser._new?"➕ Tambah User":"✏️ Edit User"}</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Username *"><input value={editUser.username} onChange={e=>setEditUser(p=>({...p,username:e.target.value}))} style={inp} placeholder="username..."/></F>
            <F label="Nama Lengkap *"><input value={editUser.name} onChange={e=>setEditUser(p=>({...p,name:e.target.value}))} style={inp}/></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Role"><select value={editUser.role} onChange={e=>setEditUser(p=>({...p,role:e.target.value}))} style={sel}>{ROLES.map(r=><option key={r.id} value={r.id}>{r.label}</option>)}</select></F>
            <F label="Email"><input type="email" value={editUser.email} onChange={e=>setEditUser(p=>({...p,email:e.target.value}))} style={inp}/></F>
          </div>
          {editUser._new&&<F label="Password (default: username)"><input type="text" value={editUser.password||""} onChange={e=>setEditUser(p=>({...p,password:e.target.value}))} style={inp} placeholder="Kosongkan = pakai username"/></F>}
          <F label="Status"><select value={editUser.active?"aktif":"nonaktif"} onChange={e=>setEditUser(p=>({...p,active:e.target.value==="aktif"}))} style={sel}><option value="aktif">Aktif</option><option value="nonaktif">Non-Aktif</option></select></F>
          {editUser.role&&<div style={{background:C.surface,borderRadius:8,padding:12,fontSize:12}}>
            <div style={{fontWeight:600,marginBottom:6}}>Akses Modul:</div>
            <div style={{color:C.muted}}>{ROLES.find(r=>r.id===editUser.role)?.pages[0]==="all"?"Semua modul":ROLES.find(r=>r.id===editUser.role)?.pages.join(", ")}</div>
          </div>}
          <button onClick={()=>saveUser(editUser)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}
      {confirmDel&&<ConfirmModal msg={`Hapus user "${users.find(u=>u.id===confirmDel)?.username}"?`} onConfirm={()=>deleteUser(confirmDel)} onCancel={()=>setConfirmDel(null)}/>}
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function ERPSystem(){
  const [page,setPage]=useState("dashboard");
  const [products,setProducts]=useState(seedProducts);
  const [customers,setCustomers]=useState(seedCustomers);
  const [orders,setOrders]=useState(seedOrders);
  const [invoices,setInvoices]=useState(seedInvoices);
  const [movements,setMovements]=useState(seedMovements);
  const [manualCF,setManualCF]=useState(seedCashFlows);
  const [employees,setEmployees]=useState(seedEmployees);
  const [payrolls,setPayrolls]=useState(seedPayrolls);
  const [suppliers,setSuppliers]=useState(seedSuppliers);
  const [pos,setPOs]=useState(seedPOs);
  const [returns,setReturns]=useState(seedReturns);
  const [warehouses,setWarehouses]=useState(seedWarehouses);
  const [discounts,setDiscounts]=useState(seedDiscounts);
  const [attendance,setAttendance]=useState(seedAttendance);
  const [salesData]=useState(seedSalesChart);
  const [users,setUsers]=useState(seedUsers);
  const [currentUser,setCurrentUser]=useState(seedUsers[0]);
  const [showNotif,setShowNotif]=useState(false);
  // modals
  const [editOrder,setEditOrder]=useState(null);
  const [confirmDelOrder,setConfirmDelOrder]=useState(null);
  const [detailOrder,setDetailOrder]=useState(null);
  const [editInvoice,setEditInvoice]=useState(null);
  const [confirmDelInv,setConfirmDelInv]=useState(null);
  const [editCustomer,setEditCustomer]=useState(null);
  const [confirmDelCust,setConfirmDelCust]=useState(null);
  const [invoiceCreated,setInvoiceCreated]=useState(null);

  const nav=[
    {id:"dashboard",icon:"📈",label:"Dashboard"},
    {id:"inventory",icon:"📦",label:"Inventori"},
    {id:"orders",icon:"🛒",label:"Pesanan"},
    {id:"invoices",icon:"🧾",label:"Invoice"},
    {id:"cashflow",icon:"💵",label:"Arus Kas"},
    {id:"po",icon:"📋",label:"Purchase Order"},
    {id:"retur",icon:"↩️",label:"Retur"},
    {id:"gudang",icon:"🏭",label:"Multi-Gudang"},
    {id:"diskon",icon:"🏷️",label:"Diskon"},
    {id:"hr",icon:"👥",label:"HR & Gaji"},
    {id:"absensi",icon:"📅",label:"Absensi"},
    {id:"laporan",icon:"📊",label:"Laporan"},
    {id:"customers",icon:"🏢",label:"Pelanggan"},
    {id:"users",icon:"🔐",label:"User & Role"},
  ];

  const lowStock=products.filter(p=>p.stock<=p.minStock);
  const notifCount=useMemo(()=>{
    let n=lowStock.length;
    const todayD=new Date(today());
    invoices.filter(i=>i.status!=="paid").forEach(i=>{if(Math.ceil((new Date(i.dueDate)-todayD)/(1000*60*60*24))<=3)n++;});
    pos.filter(p=>p.status==="dipesan"&&Math.ceil((new Date(today())-new Date(p.date))/(1000*60*60*24))>7).forEach(()=>n++);
    return n;
  },[lowStock,invoices,pos]);
  const allCF=useMemo(()=>{
    const fi=invoices.filter(i=>i.status==="paid"&&i.paidDate).map(i=>({id:`CF-INV-${i.id}`,date:i.paidDate,type:"in",category:"Penjualan",description:`${i.id} – ${i.customer}`,amount:i.grandTotal,source:"invoice"}));
    const fp=payrolls.filter(p=>p.status==="dibayar"&&p.payDate).map(p=>({id:`CF-PAY-${p.id}`,date:p.payDate,type:"out",category:"Gaji",description:`Gaji ${p.employeeName} – ${p.month}`,amount:p.takehome,ref:p.id,source:"payroll"}));
    return[...fi,...fp,...manualCF.map(m=>({...m,source:"manual"}))].sort((a,b)=>a.date.localeCompare(b.date));
  },[invoices,manualCF,payrolls]);
  const netCash=allCF.reduce((s,e)=>s+(e.type==="in"?e.amount:-e.amount),0);

  // ── Orders CRUD ──
  const saveOrder=(o)=>{
    if(o._new){const no={...o,id:`SO-${Date.now()}`,_new:undefined};setOrders(prev=>[...prev,no]);}
    else{setOrders(prev=>prev.map(x=>x.id===o.id?o:x));}
    setEditOrder(null);
  };
  const deleteOrder=(id)=>{
    const o=orders.find(x=>x.id===id);
    if(o?.invoiceId)setInvoices(prev=>prev.filter(i=>i.id!==o.invoiceId));
    setOrders(prev=>prev.filter(x=>x.id!==id));
    setConfirmDelOrder(null);
  };

  // ── Invoice CRUD ──
  const createInvoice=(order)=>{
    const tax=order.total*0.1;
    const inv={id:`INV-${Date.now()}`,orderId:order.id,date:today(),dueDate:new Date(Date.now()+14*86400000).toISOString().slice(0,10),customer:order.customer,total:order.total,tax,grandTotal:order.total+tax,status:"unpaid",paidDate:null};
    setInvoices(prev=>[...prev,inv]);
    setOrders(prev=>prev.map(o=>o.id===order.id?{...o,invoiceId:inv.id,status:"delivered"}:o));
    setInvoiceCreated(inv);
  };
  const saveInvoice=(inv)=>{setInvoices(prev=>prev.map(x=>x.id===inv.id?inv:x));setEditInvoice(null);};
  const deleteInvoice=(id)=>{
    const inv=invoices.find(i=>i.id===id);
    if(inv?.orderId)setOrders(prev=>prev.map(o=>o.id===inv.orderId?{...o,invoiceId:null}:o));
    setInvoices(prev=>prev.filter(i=>i.id!==id));
    setConfirmDelInv(null);
  };
  const markPaid=(id)=>setInvoices(prev=>prev.map(i=>i.id===id?{...i,status:"paid",paidDate:today()}:i));

  // ── Customers CRUD ──
  const saveCustomer=(c)=>{
    if(c._new){setCustomers(prev=>[...prev,{...c,id:uid(),_new:undefined}]);}
    else{setCustomers(prev=>prev.map(x=>x.id===c.id?c:x));}
    setEditCustomer(null);
  };
  const deleteCustomer=(id)=>{setCustomers(prev=>prev.filter(c=>c.id!==id));setConfirmDelCust(null);};

  const thS={padding:"9px 12px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`2px solid ${C.border}`};

  return(
    <div style={{display:"flex",minHeight:"100vh",fontFamily:"'Inter',system-ui,sans-serif",background:C.bg,color:C.text}}>
      {/* SIDEBAR */}
      <aside style={{width:220,background:C.sidebar,display:"flex",flexDirection:"column",padding:"0 0 24px",flexShrink:0}}>
        <div style={{padding:"20px 16px 16px",borderBottom:"1px solid rgba(255,255,255,0.07)",background:"linear-gradient(135deg,rgba(0,212,255,0.06),rgba(255,0,255,0.04))"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{width:38,height:38,borderRadius:9,background:"#0A1628",border:"2px solid rgba(0,212,255,0.35)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,boxShadow:"0 0 12px rgba(0,212,255,0.15)"}}>
              <span style={{fontWeight:900,fontSize:13,letterSpacing:"-1px",lineHeight:1}}>
                <span style={{color:"#00D4FF"}}>G</span><span style={{color:"#FF00FF"}}>K</span><span style={{color:"#FFD700"}}>S</span>
              </span>
            </div>
            <div>
              <div style={{fontWeight:700,fontSize:11,color:"#FFF",lineHeight:1.3}}>Global Karunia Supply</div>
              <div style={{fontSize:9,color:"#64748B",marginTop:1}}>Digital Printing Distributor</div>
            </div>
          </div>
          <div style={{height:2,borderRadius:2,background:"linear-gradient(90deg,#00D4FF,#FF00FF,#FFD700)"}}/>
        </div>
        <nav style={{flex:1,padding:"12px 10px"}}>
          {nav.map(n=>(
            <button key={n.id} onClick={()=>setPage(n.id)} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"10px 12px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:500,background:page===n.id?"rgba(0,180,216,0.85)":"transparent",color:page===n.id?"#FFF":"#94A3B8",marginBottom:2,transition:"all 0.15s"}}>
              <span style={{fontSize:15}}>{n.icon}</span>{n.label}
              {n.id==="inventory"&&lowStock.length>0&&<span style={{marginLeft:"auto",background:C.danger,color:"#FFF",borderRadius:10,fontSize:10,padding:"1px 6px",fontWeight:700}}>{lowStock.length}</span>}
            </button>
          ))}
        </nav>
        <div style={{padding:"0 12px"}}>
          <div style={{background:"rgba(0,212,255,0.08)",borderRadius:10,padding:12,border:"1px solid rgba(0,212,255,0.2)",marginBottom:10}}>
            <div style={{fontSize:11,color:"#64748B",marginBottom:4}}>Saldo Kas Bersih</div>
            <div style={{color:netCash>=0?C.success:C.danger,fontWeight:700,fontSize:16}}>{fmt(netCash)}</div>
          </div>
          <div style={{background:"rgba(255,215,0,0.08)",borderRadius:10,padding:12,border:"1px solid rgba(255,215,0,0.2)"}}>
            <div style={{fontSize:11,color:"#94A3B8",marginBottom:4}}>Stok Kritis</div>
            <div style={{color:C.danger,fontWeight:700,fontSize:18}}>{lowStock.length}</div>
          </div>
        </div>
      </aside>

      {/* MAIN */}
      <main style={{flex:1,background:C.surface,overflow:"auto"}}>
        <div style={{background:"#FFF",borderBottom:`1px solid ${C.border}`,padding:"12px 28px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            {/* Mini GKS badge in topbar */}
            <div style={{display:"flex",gap:2,fontWeight:900,fontSize:16,letterSpacing:"-0.5px"}}>
              <span style={{color:"#00D4FF"}}>G</span><span style={{color:"#FF00FF"}}>K</span><span style={{color:"#FFD700"}}>S</span>
            </div>
            <div style={{width:1,height:28,background:C.border}}/>
            <div>
              <div style={{fontWeight:700,fontSize:17,color:C.text}}>{nav.find(n=>n.id===page)?.label}</div>
              <div style={{fontSize:11,color:C.muted}}>{new Date().toLocaleDateString("id-ID",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            {/* Notif Bell */}
            <div style={{position:"relative"}}>
              <button onClick={()=>setShowNotif(v=>!v)} style={{background:showNotif?C.primaryL:"#FFF",border:`1px solid ${C.border}`,borderRadius:8,padding:"7px 12px",cursor:"pointer",fontSize:16,position:"relative"}}>
                🔔
                {notifCount>0&&<span style={{position:"absolute",top:-4,right:-4,background:C.danger,color:"#FFF",borderRadius:"50%",width:18,height:18,fontSize:10,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>{notifCount}</span>}
              </button>
              {showNotif&&<NotifPanel products={products} invoices={invoices} pos={pos} onClose={()=>setShowNotif(false)}/>}
            </div>
            {/* Current User */}
            {currentUser&&(
              <div style={{display:"flex",alignItems:"center",gap:8,background:C.surface,borderRadius:8,padding:"5px 12px",border:`1px solid ${C.border}`,cursor:"pointer"}} onClick={()=>setPage("users")}>
                <div style={{width:26,height:26,borderRadius:"50%",background:C.primary,color:"#FFF",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:12}}>{currentUser.name[0]}</div>
                <div><div style={{fontSize:12,fontWeight:600,lineHeight:1.2}}>{currentUser.name}</div><div style={{fontSize:10,color:C.muted}}><Badge status={currentUser.role}/></div></div>
              </div>
            )}
            {page==="orders"&&<button onClick={()=>setEditOrder({_new:true,date:today(),customer:"",status:"pending",items:[],total:0,invoiceId:null})} style={btn(C.primary)}>+ Pesanan Baru</button>}
            {page==="customers"&&<button onClick={()=>setEditCustomer({_new:true,name:"",contact:"",phone:"",email:"",address:""})} style={btn(C.primary)}>+ Tambah Pelanggan</button>}
            {page==="dashboard"&&<ExportBtn
              onExcel={()=>exportExcel([buildInvExport(products),buildMovExport(movements),buildOrdExport(orders),buildInvoExport(invoices),buildCFExport(allCF),buildCustExport(customers,orders)].map((e,i)=>({name:["Inventori","Mutasi","Pesanan","Invoice","Arus Kas","Pelanggan"][i],...e})))}
              onPdf={()=>{const h=[buildInvExport(products),buildOrdExport(orders),buildInvoExport(invoices),buildCustExport(customers,orders)].map(e=>e.html).join("");exportPDF("Laporan Lengkap",h);}}
            />}
          </div>
        </div>

        <div style={{padding:"24px 28px"}}>
          {/* DASHBOARD */}
          {page==="dashboard"&&<DashboardPage salesData={salesData} products={products} orders={orders} invoices={invoices} cashFlowEntries={allCF} employees={employees} payrolls={payrolls}/>}

          {/* ABSENSI */}
          {page==="absensi"&&<AbsensiPage attendance={attendance} setAttendance={setAttendance} employees={employees} payrolls={payrolls} setPayrolls={setPayrolls}/>}

          {/* USER & ROLE */}
          {page==="users"&&<UserRolePage users={users} setUsers={setUsers} currentUser={currentUser} setCurrentUser={setCurrentUser}/>}

          {/* INVENTORY */}
          {page==="inventory"&&<InventoryPage products={products} setProducts={setProducts} movements={movements} setMovements={setMovements}/>}

          {/* ORDERS */}
          {page==="orders"&&(
            <div style={card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                <div style={{fontWeight:700,fontSize:15}}>Daftar Pesanan</div>
                <ExportBtn onExcel={()=>{const e=buildOrdExport(orders);exportExcel([{name:"Pesanan",...e}]);}} onPdf={()=>{const e=buildOrdExport(orders);exportPDF("Pesanan",e.html);}}/>
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:C.surface}}>{["No. Pesanan","Tanggal","Pelanggan","Total","Status","Invoice","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
                <tbody>
                  {orders.map(o=>(
                    <tr key={o.id} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td style={tdS}><span style={{color:C.primary,fontWeight:700}}>{o.id}</span></td>
                      <td style={tdS}>{o.date}</td>
                      <td style={tdS}><div style={{fontWeight:600}}>{o.customer}</div></td>
                      <td style={tdS}><span style={{fontWeight:700}}>{fmt(o.total)}</span></td>
                      <td style={tdS}><Badge status={o.status}/></td>
                      <td style={tdS}>{o.invoiceId?<span style={{color:C.success,fontWeight:600,fontSize:12}}>{o.invoiceId}</span>:<span style={{color:C.muted,fontSize:12}}>—</span>}</td>
                      <td style={tdS}>
                        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                          <button onClick={()=>setDetailOrder(o)} style={btnSm(C.primary)}>Detail</button>
                          {!o.invoiceId&&o.status!=="cancelled"&&<button onClick={()=>createInvoice(o)} style={btnSm(C.success)}>Invoice</button>}
                          <button onClick={()=>setEditOrder({...o})} style={btnSm("#64748B")}>✏️</button>
                          <button onClick={()=>setConfirmDelOrder(o.id)} style={btnSm(C.danger)}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* INVOICES */}
          {page==="invoices"&&(
            <div style={card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                <div style={{fontWeight:700,fontSize:15}}>Daftar Invoice</div>
                <ExportBtn onExcel={()=>{const e=buildInvoExport(invoices);exportExcel([{name:"Invoice",...e}]);}} onPdf={()=>{const e=buildInvoExport(invoices);exportPDF("Invoice",e.html);}}/>
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:C.surface}}>{["No. Invoice","Pesanan","Pelanggan","Tgl","Jatuh Tempo","Subtotal","PPN","Total","Status","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
                <tbody>
                  {invoices.map(inv=>(
                    <tr key={inv.id} style={{borderBottom:`1px solid ${C.border}`}}>
                      <td style={tdS}><span style={{color:C.primary,fontWeight:700}}>{inv.id}</span></td>
                      <td style={tdS}>{inv.orderId}</td>
                      <td style={tdS}><div style={{fontWeight:600}}>{inv.customer}</div></td>
                      <td style={tdS}>{inv.date}</td>
                      <td style={tdS}>{inv.dueDate}</td>
                      <td style={tdS}>{fmt(inv.total)}</td>
                      <td style={tdS}>{fmt(inv.tax)}</td>
                      <td style={tdS}><span style={{fontWeight:700}}>{fmt(inv.grandTotal)}</span></td>
                      <td style={tdS}><Badge status={inv.status}/></td>
                      <td style={tdS}>
                        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                          {inv.status!=="paid"&&<button onClick={()=>markPaid(inv.id)} style={btnSm(C.success)}>✓ Lunas</button>}
                          <button onClick={()=>setEditInvoice({...inv})} style={btnSm("#64748B")}>✏️</button>
                          <button onClick={()=>setConfirmDelInv(inv.id)} style={btnSm(C.danger)}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{marginTop:14,padding:"13px 16px",background:C.surface,borderRadius:10,display:"flex",gap:28}}>
                <div><div style={{fontSize:12,color:C.muted}}>Total Tertagih</div><div style={{fontWeight:700,fontSize:16}}>{fmt(invoices.reduce((s,i)=>s+i.grandTotal,0))}</div></div>
                <div><div style={{fontSize:12,color:C.muted}}>Sudah Dibayar</div><div style={{fontWeight:700,fontSize:16,color:C.success}}>{fmt(invoices.filter(i=>i.status==="paid").reduce((s,i)=>s+i.grandTotal,0))}</div></div>
                <div><div style={{fontSize:12,color:C.muted}}>Belum Dibayar</div><div style={{fontWeight:700,fontSize:16,color:C.danger}}>{fmt(invoices.filter(i=>i.status!=="paid").reduce((s,i)=>s+i.grandTotal,0))}</div></div>
                <div style={{marginLeft:"auto",background:C.primaryL,padding:"8px 14px",borderRadius:8}}><div style={{fontSize:11,color:C.primary,fontWeight:600}}>💡 Invoice lunas otomatis masuk ke Arus Kas</div></div>
              </div>
            </div>
          )}

          {/* CASH FLOW */}
          {page==="cashflow"&&<CashFlowPage invoices={invoices} manualCF={manualCF} setManualCF={setManualCF}/>}

          {/* PURCHASE ORDER */}
          {page==="po"&&<POPage pos={pos} setPOs={setPOs} suppliers={suppliers} setSuppliers={setSuppliers} products={products} setProducts={setProducts} setMovements={setMovements}/>}

          {/* RETUR */}
          {page==="retur"&&<ReturPage returns={returns} setReturns={setReturns} products={products} setProducts={setProducts} setMovements={setMovements} orders={orders} customers={customers} pos={pos} suppliers={suppliers}/>}

          {/* MULTI-GUDANG */}
          {page==="gudang"&&<GudangPage warehouses={warehouses} setWarehouses={setWarehouses} products={products} setProducts={setProducts} setMovements={setMovements}/>}

          {/* DISKON */}
          {page==="diskon"&&<DiskonPage discounts={discounts} setDiscounts={setDiscounts} customers={customers} products={products}/>}

          {/* HR & GAJI */}
          {page==="hr"&&<HRPage employees={employees} setEmployees={setEmployees} payrolls={payrolls} setPayrolls={setPayrolls}/>}

          {/* LAPORAN */}
          {page==="laporan"&&<LaporanPage products={products} movements={movements} orders={orders} invoices={invoices} cashFlowEntries={allCF} employees={employees} payrolls={payrolls}/>}

          {/* CUSTOMERS */}
          {page==="customers"&&(
            <div style={card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                <div style={{fontWeight:700,fontSize:15}}>Daftar Pelanggan</div>
                <ExportBtn onExcel={()=>{const e=buildCustExport(customers,orders);exportExcel([{name:"Pelanggan",...e}]);}} onPdf={()=>{const e=buildCustExport(customers,orders);exportPDF("Pelanggan",e.html);}}/>
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:C.surface}}>{["Perusahaan","PIC","Telepon","Email","Alamat","Total Pesanan","Aksi"].map(h=><th key={h} style={thS}>{h}</th>)}</tr></thead>
                <tbody>
                  {customers.map(c=>{
                    const co=orders.filter(o=>o.customer===c.name);
                    return(
                      <tr key={c.id} style={{borderBottom:`1px solid ${C.border}`}}>
                        <td style={tdS}><div style={{fontWeight:700}}>{c.name}</div></td>
                        <td style={tdS}>{c.contact}</td>
                        <td style={tdS}>{c.phone}</td>
                        <td style={tdS}><a href={`mailto:${c.email}`} style={{color:C.primary}}>{c.email}</a></td>
                        <td style={tdS}><div style={{fontSize:12,color:C.muted,maxWidth:180}}>{c.address}</div></td>
                        <td style={tdS}><div style={{fontWeight:700,color:C.primary}}>{fmt(co.reduce((s,o)=>s+o.total,0))}</div><div style={{fontSize:11,color:C.muted}}>{co.length} pesanan</div></td>
                        <td style={tdS}>
                          <div style={{display:"flex",gap:5}}>
                            <button onClick={()=>setEditCustomer({...c})} style={btnSm("#64748B")}>✏️ Edit</button>
                            <button onClick={()=>setConfirmDelCust(c.id)} style={btnSm(C.danger)}>🗑️</button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>

      {/* ── ORDER DETAIL MODAL ── */}
      {detailOrder&&<Modal onClose={()=>setDetailOrder(null)}>
        <div style={{fontWeight:700,fontSize:18,marginBottom:4}}>Detail Pesanan</div>
        <div style={{color:C.primary,fontWeight:600,marginBottom:16}}>{detailOrder.id}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
          {[["Pelanggan",detailOrder.customer],["Tanggal",detailOrder.date],["Status",<Badge status={detailOrder.status}/>],["Invoice",detailOrder.invoiceId||"—"]].map(([k,v])=>(
            <div key={k} style={{background:C.surface,padding:12,borderRadius:8}}><div style={{fontSize:11,color:C.muted,marginBottom:4}}>{k}</div><div style={{fontWeight:600}}>{v}</div></div>
          ))}
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,marginBottom:16}}>
          <thead><tr style={{background:C.surface}}>{["Produk","Qty","Harga","Subtotal"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontWeight:600,color:C.muted,fontSize:12,borderBottom:`1px solid ${C.border}`}}>{h}</th>)}</tr></thead>
          <tbody>{detailOrder.items.map((it,i)=><tr key={i} style={{borderBottom:`1px solid ${C.border}`}}><td style={tdS}>{it.name}</td><td style={tdS}>{it.qty}</td><td style={tdS}>{fmt(it.price)}</td><td style={tdS}><b>{fmt(it.qty*it.price)}</b></td></tr>)}</tbody>
        </table>
        <div style={{textAlign:"right",borderTop:`2px solid ${C.border}`,paddingTop:12}}>
          <div style={{fontSize:13,color:C.muted}}>Subtotal: {fmt(detailOrder.total)}</div>
          <div style={{fontSize:13,color:C.muted}}>PPN 10%: {fmt(detailOrder.total*0.1)}</div>
          <div style={{fontWeight:700,fontSize:17,color:C.primary,marginTop:6}}>Total: {fmt(detailOrder.total*1.1)}</div>
        </div>
        {!detailOrder.invoiceId&&detailOrder.status!=="cancelled"&&<button onClick={()=>{createInvoice(detailOrder);setDetailOrder(null);}} style={{...btn(C.success),marginTop:16,width:"100%"}}>🧾 Buat Invoice</button>}
      </Modal>}

      {/* ── EDIT ORDER MODAL ── */}
      {editOrder&&<Modal onClose={()=>setEditOrder(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editOrder._new?"➕ Pesanan Baru":"✏️ Edit Pesanan"}</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Pelanggan"><select value={editOrder.customer} onChange={e=>setEditOrder(p=>({...p,customer:e.target.value}))} style={sel}><option value="">— Pilih —</option>{customers.map(c=><option key={c.id}>{c.name}</option>)}</select></F>
            <F label="Tanggal"><input type="date" value={editOrder.date} onChange={e=>setEditOrder(p=>({...p,date:e.target.value}))} style={inp}/></F>
          </div>
          <F label="Status"><select value={editOrder.status} onChange={e=>setEditOrder(p=>({...p,status:e.target.value}))} style={sel}>{["pending","processing","delivered","cancelled"].map(s=><option key={s}>{s}</option>)}</select></F>
          <F label="Total (IDR)"><input type="number" min="0" value={editOrder.total} onChange={e=>setEditOrder(p=>({...p,total:Number(e.target.value)}))} style={inp}/></F>
          <button onClick={()=>saveOrder(editOrder)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* ── EDIT INVOICE MODAL ── */}
      {editInvoice&&<Modal onClose={()=>setEditInvoice(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>✏️ Edit Invoice</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Tanggal Invoice"><input type="date" value={editInvoice.date} onChange={e=>setEditInvoice(p=>({...p,date:e.target.value}))} style={inp}/></F>
            <F label="Jatuh Tempo"><input type="date" value={editInvoice.dueDate} onChange={e=>setEditInvoice(p=>({...p,dueDate:e.target.value}))} style={inp}/></F>
          </div>
          <F label="Status"><select value={editInvoice.status} onChange={e=>setEditInvoice(p=>({...p,status:e.target.value,paidDate:e.target.value==="paid"?(p.paidDate||today()):null}))} style={sel}><option value="unpaid">Belum Bayar</option><option value="paid">Lunas</option></select></F>
          {editInvoice.status==="paid"&&<F label="Tanggal Bayar"><input type="date" value={editInvoice.paidDate||""} onChange={e=>setEditInvoice(p=>({...p,paidDate:e.target.value}))} style={inp}/></F>}
          <button onClick={()=>saveInvoice(editInvoice)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* ── EDIT CUSTOMER MODAL ── */}
      {editCustomer&&<Modal onClose={()=>setEditCustomer(null)}>
        <div style={{fontWeight:700,fontSize:17,marginBottom:20}}>{editCustomer._new?"➕ Tambah Pelanggan":"✏️ Edit Pelanggan"}</div>
        <div style={{display:"grid",gap:13}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Nama Perusahaan *"><input value={editCustomer.name} onChange={e=>setEditCustomer(p=>({...p,name:e.target.value}))} style={inp} placeholder="PT / CV ..."/></F>
            <F label="Nama PIC"><input value={editCustomer.contact} onChange={e=>setEditCustomer(p=>({...p,contact:e.target.value}))} style={inp} placeholder="Nama kontak"/></F>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <F label="Telepon"><input value={editCustomer.phone} onChange={e=>setEditCustomer(p=>({...p,phone:e.target.value}))} style={inp} placeholder="021-..."/></F>
            <F label="Email"><input type="email" value={editCustomer.email} onChange={e=>setEditCustomer(p=>({...p,email:e.target.value}))} style={inp} placeholder="email@..."/></F>
          </div>
          <F label="Alamat"><input value={editCustomer.address} onChange={e=>setEditCustomer(p=>({...p,address:e.target.value}))} style={inp} placeholder="Jl. ..."/></F>
          <button onClick={()=>saveCustomer(editCustomer)} style={{...btn(C.primary),padding:"11px",marginTop:4}}>💾 Simpan</button>
        </div>
      </Modal>}

      {/* ── INVOICE CREATED MODAL ── */}
      {invoiceCreated&&<Modal onClose={()=>setInvoiceCreated(null)}>
        <div style={{textAlign:"center",padding:"10px 0"}}>
          <div style={{fontSize:48,marginBottom:12}}>✅</div>
          <div style={{fontWeight:700,fontSize:18,marginBottom:8}}>Invoice Berhasil Dibuat!</div>
          <div style={{color:C.primary,fontWeight:600,fontSize:16,marginBottom:16}}>{invoiceCreated.id}</div>
          <div style={{background:C.surface,borderRadius:10,padding:16,marginBottom:14,textAlign:"left"}}>
            {[["Pelanggan",invoiceCreated.customer],["Tanggal",invoiceCreated.date],["Jatuh Tempo",invoiceCreated.dueDate],["Subtotal",fmt(invoiceCreated.total)],["PPN 10%",fmt(invoiceCreated.tax)],["Total",fmt(invoiceCreated.grandTotal)]].map(([k,v])=>(
              <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.border}`}}><span style={{color:C.muted,fontSize:13}}>{k}</span><span style={{fontWeight:600,fontSize:13}}>{v}</span></div>
            ))}
          </div>
          <div style={{background:C.primaryL,borderRadius:8,padding:"10px 14px",marginBottom:14,fontSize:12,color:C.primary,fontWeight:500}}>💵 Tandai <b>Lunas</b> → otomatis masuk ke Arus Kas</div>
          <button onClick={()=>{setInvoiceCreated(null);setPage("invoices");}} style={btn(C.primary)}>Lihat di Invoice →</button>
        </div>
      </Modal>}

      {/* ── CONFIRM DEL MODALS ── */}
      {confirmDelOrder&&<ConfirmModal msg={`Hapus pesanan "${confirmDelOrder}"? Invoice terkait juga dihapus.`} onConfirm={()=>deleteOrder(confirmDelOrder)} onCancel={()=>setConfirmDelOrder(null)}/>}
      {confirmDelInv&&<ConfirmModal msg={`Hapus invoice "${confirmDelInv}"?`} onConfirm={()=>deleteInvoice(confirmDelInv)} onCancel={()=>setConfirmDelInv(null)}/>}
      {confirmDelCust&&<ConfirmModal msg={`Hapus pelanggan "${customers.find(c=>c.id===confirmDelCust)?.name}"?`} onConfirm={()=>deleteCustomer(confirmDelCust)} onCancel={()=>setConfirmDelCust(null)}/>}
    </div>
  );
}
