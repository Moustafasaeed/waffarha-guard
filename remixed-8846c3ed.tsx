// @ts-nocheck
import { useState, useRef, useCallback, useEffect } from "react";
import {
  logUploadSession, logFraudAlerts, logAuditEntry,
  buildCCAlerts, buildHighAmtAlerts, buildFakeDomAlerts,
  buildWalletAbuserAlerts, buildBNPLAlerts, buildAdminAlerts, buildFawryAlerts, buildPromoAlerts,
  loadBlockedEmails, blockEmail, unblockEmail, loadDashboardStats,
} from "./lib/supabase";
import { Shield, Upload, Search, X, CheckCircle, Eye, AlertTriangle, Users, CreditCard, DollarSign, ShieldAlert, Download } from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// ─── Auth & Users ─────────────────────────────────────────────────────────────
const SUPER_ADMIN = { username:"MSM", password:"sa3EED5544201**", role:"superadmin" };
function getSuperAdminPassword() { try { return sessionStorage.getItem("fg_sa_pw") || SUPER_ADMIN.password; } catch { return SUPER_ADMIN.password; } }
function setSuperAdminPassword(pw) { try { sessionStorage.setItem("fg_sa_pw", pw); } catch {} }
const INITIAL_USERS = [
  { username:"Yaheia.adel",   password:"asd123456**",  role:"user" },
  { username:"Traek.nabil",   password:"asd789**",     role:"user" },
  { username:"Fatma.saad",    password:"asd101112**",  role:"user" },
  { username:"Mostafa.ezzat", password:"asd131415**",  role:"user" },
];
const SECURITY_QUESTION = "What is your favorite superhero?";
function loadUsers() { try { const s=sessionStorage.getItem("fg_users"); return s?JSON.parse(s):INITIAL_USERS; } catch { return INITIAL_USERS; } }
function saveUsers(u) { try { sessionStorage.setItem("fg_users", JSON.stringify(u)); } catch {} }

// ─── Audit ────────────────────────────────────────────────────────────────────
const AUDIT_LOG={items:[]};
function addAuditLog(e){AUDIT_LOG.items.unshift({...e,time:new Date().toLocaleString(),id:Date.now()});}
function getAuditLog(){return AUDIT_LOG.items;}

// ─── Utilities ────────────────────────────────────────────────────────────────
const ALLOWED_DOMAINS=new Set(["gmail.com","hotmail.com","yahoo.com","icloud.com","outlook.com","msn.com","live.com"]);
function isDisposable(email){const d=(email.split("@")[1]||"").toLowerCase().trim();if(!d)return false;if(ALLOWED_DOMAINS.has(d))return false;if(d.endsWith(".eg")||d.endsWith(".edu"))return false;return true;}
function getCol(row){return(...names)=>{for(const n of names){const k=Object.keys(row).find(k=>k.trim().toLowerCase()===n.toLowerCase());if(k!==undefined&&row[k]!==undefined&&row[k].toString().trim()!=="")return row[k].toString().trim();}return"";};}
function formatTimestamp(val){if(!val)return"";const n=parseFloat(val.toString());if(!isNaN(n)&&n>40000&&n<60000){const d=new Date(Math.round((n-25569)*86400*1000));const p=x=>x.toString().padStart(2,"0");return`${p(d.getUTCDate())}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`;}return val.toString();}

// ─── Enrichment ───────────────────────────────────────────────────────────────
function buildCCDetails(row){const g=getCol(row),pd=g("Payment Description"),ey=g("Expiry Year"),em=g("Expiry Month");if(!pd&&!ey&&!em)return"";return`${pd} ${ey}/${em.toString().padStart(2,"0")}`.trim().replace(/\s+/g," ");}
function getPayTabsAmt(row){const keys=Object.keys(row),ci=keys.findIndex(k=>k.trim().toLowerCase()==="customer name");if(ci>0){const n=parseFloat((row[keys[ci-1]]||"").toString().replace(/[^0-9.]/g,""));if(n>0)return n;}for(const k of keys.filter(k=>/amount/i.test(k.trim()))){const n=parseFloat((row[k]||"").toString().replace(/[^0-9.]/g,""));if(n>0)return n;}return 0;}
function enrichPayTabsRow(row){const g=getCol(row);return{...row,_cc:buildCCDetails(row),_amt:getPayTabsAmt(row),_email:g("Customer Email","Email","email").toLowerCase().trim(),_country:g("Issuer Country"),_name:g("Customer Name"),_cartId:g("Cart ID","CartID","cart id"),_timestamp:formatTimestamp(g("Timestamp","timestamp","TIMESTAMP")),_authMsg:g("Auth Message","Auth_Message","auth message"),_eci:g("ECI","eci"),_orderstatus:""};}

function mergeNoonAdmin(noon,admin){const lk={};admin.forEach(r=>{const k=getCol(r)("Transaction ID","transaction id");if(k)lk[k.trim()]=r;});return noon.map(nr=>{const k=getCol(nr)("Merchantorderreference","merchantorderreference");const ar=k?(lk[k.trim()]||{}):{}; return{...ar,...nr};});}
function getNoonAmt(row){const g=getCol(row),raw=g("Amount","Salesamount","salesamount","Capturedamount","capturedamount","Authorizedamount","authorizedamount");return parseFloat((raw||"").toString().replace(/[^0-9.]/g,""))||0;}
function enrichNoonRow(row){const g=getCol(row),fn=g("Firstname","firstname"),ln=g("Lastname","lastname");return{...row,_cc:g("Payerinfo","payerinfo"),_amt:getNoonAmt(row),_email:(g("User Email","useremail","Email","email")).toLowerCase().trim(),_country:g("Issuercountry","issuercountry"),_name:g("User Name","username","Ordername","ordername")||[fn,ln].filter(Boolean).join(" "),_cartId:g("Merchantorderreference","merchantorderreference"),_timestamp:formatTimestamp(g("Orderdate_UTC","orderdate_utc","Transaction create time")),_authMsg:g("Responsemessage","responsemessage"),_eci:g("ECI","eci"),_orderstatus:g("Orderstatus","orderstatus")};}

function getPaymobAmt(row){const g=getCol(row),raw=g("amount_whole","amount");return parseFloat((raw||"").toString().replace(/[^0-9.]/g,""))||0;}
function enrichPaymobRow(row){const g=getCol(row);return{...row,_cc:g("client_phone"),_amt:getPaymobAmt(row),_email:g("client_email","email").toLowerCase().trim(),_country:"",_name:g("client_name"),_cartId:g("merchant_order_id","order_id"),_timestamp:formatTimestamp(g("created_at","Created At","timestamp","date")),_authMsg:g("data_message_execl","data_message","message","auth_message"),_eci:"",_orderstatus:g("success","Success","is_success","is_paid")};}
function enrichBNPLRow(row){const g=getCol(row);return{...row,_cc:g("payment_method"),_amt:getPaymobAmt(row),_email:g("client_email","email").toLowerCase().trim(),_country:"",_name:g("client_name"),_cartId:g("merchant_order_id","order_id"),_timestamp:formatTimestamp(g("created_at","Created At","timestamp","date")),_authMsg:g("data_message_execl","data_message","message"),_eci:"",_orderstatus:g("success","Success","is_success","is_paid")};}
const BNPL_EXCLUDE=["bank installment","mobile wallet"];
function filterBNPLRows(rows){return rows.filter(r=>{const pm=getCol(r)("payment_method").toLowerCase().trim();return!BNPL_EXCLUDE.includes(pm);});}

function parseFawryUser(raw){const parts=(raw||"").split(" - ").map(s=>s.trim());return{_userId:parts[0]||"",_phone:parts[1]||"",_name:parts.slice(2).join(" - ")||""};}
function enrichFawryRow(row){const g=getCol(row);const raw=g("user","User");const{_userId,_phone,_name}=parseFawryUser(raw);const amt=parseFloat((g("Requested","requested amount","requested")||"").replace(/[^0-9.]/g,""))||0;return{...row,_userId,_phone,_name:_name||g("Name","name","customer name")||"",_email:g("user email","user_email","email").toLowerCase().trim(),_txId:g("id","ID"),_amt:amt,_fawryCode:g("Fawry Code","fawry code","fawry_code"),_status:g("Status","status"),_payMethod:g("Payment Method","payment method"),_soldDate:formatTimestamp(g("Sold Date","sold date")),_paidDate:formatTimestamp(g("Paid Date","paid date"))};}

const TRIAL_EXCLUDE=["paid","refunded","refund"];
function enrichAdminRow(row){const g=getCol(row);const status=(g("Status")||"").toLowerCase().trim();const isTrial=!TRIAL_EXCLUDE.includes(status);const amt=parseFloat((g("Requested Amount")||"").replace(/[^0-9.]/g,""))||0;return{...row,_email:g("User Email","email").toLowerCase().trim(),_userId:g("User Id","User ID"),_name:g("User Name"),_payMethod:g("Payment Method"),_biller:g("Service Biller","Service biller"),_status:g("Status"),_amt:amt,_txId:g("Transaction ID","Transaction Id"),_timestamp:formatTimestamp(g("Transaction create time")),_isTrial:isTrial};}

function enrichPromoRow(row){const g=getCol(row);const disc=parseFloat((g("Discount Amount")||"").toString().replace(/[^0-9.]/g,""))||0;const total=parseFloat((g("total")||"").toString().replace(/[^0-9.]/g,""))||0;const req=parseFloat((g("Requested")||"").toString().replace(/[^0-9.]/g,""))||0;return{...row,_email:g("User Email","email").toLowerCase().trim(),_userId:g("User Id","User ID"),_name:g("User Name"),_mobile:g("User Mobile"),_orderId:g("Order Id","Order ID"),_promoCode:g("Promo Code"),_discountAmt:disc,_cardNum:g("Card Number"),_walletNum:g("Wallet Num"),_total:total,_requested:req,_orderStatus:g("Order Status"),_paymentName:g("Payment Name"),_paymentProvider:g("Payment Provider"),_createdAt:formatTimestamp(g("Created At")),_userAccStatus:g("User Acc Status"),_userFlag:g("user_flag"),_trxFlag:g("Trx Flag","trx_flag"),_rrn:g("RRN")};}

// ─── Detection ────────────────────────────────────────────────────────────────
function groupByEmail(enriched){const m={};enriched.forEach(r=>{if(!r._email||r._email==="n/a"||r._email==="-")return;if(!m[r._email])m[r._email]=[];m[r._email].push(r);});return m;}
function buildEmailResult(email,rows){
  const uniqueCCs=[...new Set(rows.map(r=>r._cc).filter(Boolean))];
  const custNames=[...new Set(rows.map(r=>r._name).filter(Boolean))];
  const totalAmt=rows.reduce((s,r)=>s+r._amt,0);
  const ccBreakdown={},ccCartIds={},ccTimestamps={},ccAuthMsgs={},ccEci={},ccOrderStatus={},ccTxRows={},ccCountries={};
  rows.forEach(r=>{const cc=r._cc;if(!cc)return;ccBreakdown[cc]=(ccBreakdown[cc]||0)+r._amt;if(!ccCartIds[cc])ccCartIds[cc]=new Set();if(!ccTimestamps[cc])ccTimestamps[cc]=new Set();if(!ccAuthMsgs[cc])ccAuthMsgs[cc]=new Set();if(!ccEci[cc])ccEci[cc]=new Set();if(!ccOrderStatus[cc])ccOrderStatus[cc]=new Set();if(!ccCountries[cc])ccCountries[cc]=new Set();if(!ccTxRows[cc])ccTxRows[cc]=[];if(r._cartId)ccCartIds[cc].add(r._cartId);if(r._timestamp)ccTimestamps[cc].add(r._timestamp);if(r._authMsg)ccAuthMsgs[cc].add(r._authMsg);if(r._eci)ccEci[cc].add(r._eci);if(r._orderstatus)ccOrderStatus[cc].add(r._orderstatus);if(r._country)ccCountries[cc].add(r._country);ccTxRows[cc].push({cartId:r._cartId||"",timestamp:r._timestamp||"",authMsg:r._authMsg||"",eci:r._eci||"",orderStatus:r._orderstatus||"",country:r._country||"",amt:r._amt||0});});
  [ccCartIds,ccTimestamps,ccAuthMsgs,ccEci,ccOrderStatus,ccCountries].forEach(o=>Object.keys(o).forEach(cc=>{o[cc]=[...o[cc]];}));
  const countries=[...new Set(rows.map(r=>r._country).filter(Boolean))];
  return{email,custNames,uniqueCCs,ccCartIds,ccTimestamps,ccAuthMsgs,ccEci,ccOrderStatus,ccTxRows,ccCountries,countries,totalAmt,ccBreakdown,txCount:rows.length};
}
function detectFraud(enriched,label="CC Details"){const byEmail=groupByEmail(enriched),results=[];Object.entries(byEmail).forEach(([email,rows])=>{const base=buildEmailResult(email,rows),n=base.uniqueCCs.length;const hr=[],mr=[];if(n>3)hr.push(`Used ${n} different ${label} (threshold: > 3)`);if(n===3)mr.push(`Used exactly 3 different ${label}`);let risk=null,reasons=[];if(hr.length){risk="High";reasons=hr;}else if(mr.length){risk="Mid";reasons=mr;}if(risk)results.push({...base,risk,reasons,disposable:false});});return results.sort((a,b)=>a.risk!==b.risk?(a.risk==="High"?-1:1):b.totalAmt-a.totalAmt);}
function detectHighAmounts(enriched,threshold=2000){const byEmail=groupByEmail(enriched),results=[];Object.entries(byEmail).forEach(([email,rows])=>{const hr=rows.filter(r=>r._amt>=threshold);if(!hr.length)return;const base=buildEmailResult(email,rows),max=Math.max(...hr.map(r=>r._amt));const hCCs=[...new Set(hr.map(r=>r._cc).filter(Boolean))];results.push({...base,uniqueCCs:hCCs.length?hCCs:base.uniqueCCs,highTxCount:hr.length,maxAmt:max,reasons:[`${hr.length} transaction(s) ≥ EGP ${threshold.toLocaleString()} · Highest: EGP ${max.toLocaleString()}`]});});return results.sort((a,b)=>b.maxAmt-a.maxAmt);}
function detectFakeDomain(enriched){const byEmail=groupByEmail(enriched),results=[];Object.entries(byEmail).forEach(([email,rows])=>{if(!isDisposable(email))return;const base=buildEmailResult(email,rows),domain=email.split("@")[1]||"";results.push({...base,reasons:[`Non-whitelisted email domain: @${domain}`],domain});});return results.sort((a,b)=>b.totalAmt-a.totalAmt);}
function detectBNPLFraud(enriched){const byEmail=groupByEmail(enriched),results=[];Object.entries(byEmail).forEach(([email,rows])=>{const base=buildEmailResult(email,rows),reasons=[];const failed=rows.filter(r=>(r._orderstatus||"").toString().toLowerCase().trim()==="false");if(failed.length>=3)reasons.push(`${failed.length} failed BNPL payment attempts (success = FALSE)`);const providers=[...new Set(rows.map(r=>r._cc).filter(Boolean))];if(providers.length>1)reasons.push(`Tried ${providers.length} different BNPL providers: ${providers.join(", ")}`);if(reasons.length)results.push({...base,risk:"HighSuspicious",reasons,disposable:isDisposable(email)});});return results.sort((a,b)=>b.totalAmt-a.totalAmt);}
function detectWalletAbusers(enriched){const byWallet={};enriched.forEach(r=>{const wallet=r._cc;if(!wallet)return;if(!byWallet[wallet])byWallet[wallet]={emails:new Set(),rows:[]};if(r._email&&r._email!=="n/a"&&r._email!=="-")byWallet[wallet].emails.add(r._email);byWallet[wallet].rows.push(r);});const results=[];Object.entries(byWallet).forEach(([wallet,{emails,rows}])=>{const n=emails.size;if(n<2)return;const totalAmt=rows.reduce((s,r)=>s+r._amt,0),emailList=[...emails],risk=n>=3?"High":"Mid";const reasons=[`Wallet used by ${n} different email${n>1?"s":""}: ${emailList.join(", ")}`];const byEmail={};rows.forEach(r=>{if(!r._email)return;if(!byEmail[r._email])byEmail[r._email]={cartIds:new Set(),timestamps:new Set(),authMsgs:new Set(),orderStatuses:new Set(),amt:0,txRows:[]};if(r._cartId)byEmail[r._email].cartIds.add(r._cartId);if(r._timestamp)byEmail[r._email].timestamps.add(r._timestamp);if(r._authMsg)byEmail[r._email].authMsgs.add(r._authMsg);if(r._orderstatus)byEmail[r._email].orderStatuses.add(r._orderstatus);byEmail[r._email].amt+=r._amt;byEmail[r._email].txRows.push({cartId:r._cartId||"",timestamp:r._timestamp||"",authMsg:r._authMsg||"",orderStatus:r._orderstatus||"",amt:r._amt||0});});results.push({wallet,emails:emailList,emailDetails:byEmail,totalAmt,txCount:rows.length,risk,reasons,rows});});return results.sort((a,b)=>b.emails.length-a.emails.length||b.totalAmt-a.totalAmt);}

function detectFawryHighAmt(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a"||r._email==="-")return;if(r._amt<2000)return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names}])=>{const totalAmt=rows.reduce((s,r)=>s+r._amt,0),max=Math.max(...rows.map(r=>r._amt));results.push({email,custNames:[...names],txRows:rows.map(r=>({txId:r._txId,fawryCode:r._fawryCode,status:r._status,soldDate:r._soldDate,amt:r._amt})),txCount:rows.length,totalAmt,maxAmt:max,reasons:[`${rows.length} transaction(s) ≥ EGP 2,000 · Highest: EGP ${max.toLocaleString()}`]});});return results.sort((a,b)=>b.maxAmt-a.maxAmt);}
function detectFawrySuspected(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a"||r._email==="-")return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names}])=>{const txIds=[...new Set(rows.map(r=>r._txId).filter(Boolean))];if(txIds.length<3)return;const totalAmt=rows.reduce((s,r)=>s+r._amt,0);results.push({email,custNames:[...names],txRows:rows.map(r=>({txId:r._txId,fawryCode:r._fawryCode,status:r._status,soldDate:r._soldDate,amt:r._amt})),txIds,txCount:txIds.length,totalAmt,reasons:[`${txIds.length} transactions by the same email`]});});return results.sort((a,b)=>b.txCount-a.txCount);}
function detectFawryFakeDomain(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a"||r._email==="-")return;if(!isDisposable(r._email))return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names}])=>{const totalAmt=rows.reduce((s,r)=>s+r._amt,0),domain=email.split("@")[1]||"";results.push({email,custNames:[...names],domain,txRows:rows.map(r=>({txId:r._txId,fawryCode:r._fawryCode,status:r._status,soldDate:r._soldDate,amt:r._amt})),txCount:rows.length,totalAmt,reasons:[`Non-whitelisted email domain: @${domain}`]});});return results.sort((a,b)=>b.totalAmt-a.totalAmt);}

function detectAdminPayMethods(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a")return;if(!byEmail[r._email])byEmail[r._email]={rows:[],methods:new Set(),names:new Set(),userIds:new Set()};byEmail[r._email].rows.push(r);if(r._payMethod)byEmail[r._email].methods.add(r._payMethod);if(r._name)byEmail[r._email].names.add(r._name);if(r._userId)byEmail[r._email].userIds.add(r._userId);});const results=[];Object.entries(byEmail).forEach(([email,{rows,methods,names,userIds}])=>{const n=methods.size;let risk=null,reasons=[];if(n>3){risk="High";reasons=[`Used ${n} different payment methods (threshold: > 3)`];}else if(n===3){risk="Mid";reasons=["Used exactly 3 different payment methods"];}if(risk){const totalAmt=rows.reduce((s,r)=>s+r._amt,0);const txRows=rows.map(r=>({txId:r._txId||"",status:r._status||"",amt:r._amt||0}));results.push({email,custNames:[...names],userIds:[...userIds],uniqueMethods:[...methods],txRows,txIds:rows.map(r=>r._txId).filter(Boolean),txCount:rows.length,totalAmt,risk,reasons,rows});}});return results.sort((a,b)=>a.risk!==b.risk?(a.risk==="High"?-1:1):b.totalAmt-a.totalAmt);}
function detectAdminSuspected(enriched){const byUser={};enriched.forEach(r=>{const uid=r._userId;if(!uid)return;if(!byUser[uid])byUser[uid]={rows:[],billers:new Set(),names:new Set(),emails:new Set()};byUser[uid].rows.push(r);if(r._biller)byUser[uid].billers.add(r._biller);if(r._name)byUser[uid].names.add(r._name);if(r._email)byUser[uid].emails.add(r._email);});const results=[];Object.entries(byUser).forEach(([userId,{rows,billers,names,emails}])=>{const trials=rows.filter(r=>r._isTrial);if(trials.length<5)return;const totalAmt=rows.reduce((s,r)=>s+r._amt,0);const trialBillers=[...new Set(trials.map(r=>r._biller).filter(Boolean))];const email=[...emails][0]||"";results.push({userId,email,custNames:[...names],emails:[...emails],billers:[...billers],trialBillers,trialCount:trials.length,txCount:rows.length,txRows:rows.map(r=>({txId:r._txId||"",status:r._status||"",amt:r._amt||0})),totalAmt,risk:"HighSuspicious",reasons:[`${trials.length} failed/trial attempts across ${trialBillers.length} biller(s)`],rows});});return results.sort((a,b)=>b.trialCount-a.trialCount);}
function detectAdminHighAmt(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a")return;if(r._amt<3000)return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set(),methods:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);if(r._payMethod)byEmail[r._email].methods.add(r._payMethod);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names,methods}])=>{const totalAmt=rows.reduce((s,r)=>s+r._amt,0),max=Math.max(...rows.map(r=>r._amt));results.push({email,custNames:[...names],uniqueMethods:[...methods],txRows:rows.map(r=>({txId:r._txId||"",status:r._status||"",amt:r._amt||0})),txCount:rows.length,totalAmt,maxAmt:max,risk:"HighAmount",reasons:[`${rows.length} transaction(s) ≥ EGP 3,000 · Highest: EGP ${max.toLocaleString()}`],rows});});return results.sort((a,b)=>b.maxAmt-a.maxAmt);}
function detectAdminFakeDomain(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||!isDisposable(r._email))return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set(),methods:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);if(r._payMethod)byEmail[r._email].methods.add(r._payMethod);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names,methods}])=>{const totalAmt=rows.reduce((s,r)=>s+r._amt,0),domain=email.split("@")[1]||"";results.push({email,custNames:[...names],uniqueMethods:[...methods],txRows:rows.map(r=>({txId:r._txId||"",status:r._status||"",amt:r._amt||0})),txCount:rows.length,totalAmt,risk:"FakeDomain",reasons:[`Non-whitelisted email domain: @${domain}`],domain,rows});});return results.sort((a,b)=>b.totalAmt-a.totalAmt);}
function detectRechargeAbusers(enriched){const byRecharge={};enriched.forEach(r=>{const rn=getCol(r)("recharge number","Recharge Number","recharge_number");if(!rn)return;if(!byRecharge[rn])byRecharge[rn]={emails:new Set(),rows:[]};if(r._email&&r._email!=="n/a"&&r._email!=="-")byRecharge[rn].emails.add(r._email);byRecharge[rn].rows.push(r);});const results=[];Object.entries(byRecharge).forEach(([recharge,{emails,rows}])=>{const n=emails.size;if(n<2)return;const totalAmt=rows.reduce((s,r)=>s+r._amt,0),emailList=[...emails],risk=n>=3?"High":"Mid";const byEmail={};rows.forEach(r=>{if(!r._email)return;if(!byEmail[r._email])byEmail[r._email]={txIds:new Set(),timestamps:new Set(),billers:new Set(),statuses:new Set(),amt:0,txRows:[]};if(r._txId)byEmail[r._email].txIds.add(r._txId);if(r._timestamp)byEmail[r._email].timestamps.add(r._timestamp);if(r._biller)byEmail[r._email].billers.add(r._biller);if(r._status)byEmail[r._email].statuses.add(r._status);byEmail[r._email].amt+=r._amt;byEmail[r._email].txRows.push({txId:r._txId||"",timestamp:r._timestamp||"",biller:r._biller||"",status:r._status||"",amt:r._amt||0});});results.push({recharge,emails:emailList,emailDetails:byEmail,totalAmt,txCount:rows.length,risk,reasons:[`Recharge number ${recharge} topped up by ${n} different email${n>1?"s":""}: ${emailList.join(", ")}`],rows});});return results.sort((a,b)=>b.emails.length-a.emails.length||b.totalAmt-a.totalAmt);}

function detectPromoHighDiscount(enriched,threshold=200){const byEmail={};enriched.forEach(r=>{if(!r._email||r._email==="n/a"||r._email==="-")return;if(!r._promoCode)return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set(),codes:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);if(r._promoCode)byEmail[r._email].codes.add(r._promoCode);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names,codes}])=>{const totalDiscount=rows.reduce((s,r)=>s+r._discountAmt,0);if(totalDiscount<threshold)return;results.push({email,custNames:[...names],promoCodes:[...codes],totalDiscount,txCount:rows.length,txRows:rows.map(r=>({orderId:r._orderId,promoCode:r._promoCode,trxFlag:r._trxFlag||"",discountAmt:r._discountAmt,orderStatus:r._orderStatus,total:r._total,createdAt:r._createdAt})),reasons:[`Total discount EGP ${totalDiscount.toLocaleString()} across ${rows.length} order(s)`]});});return results.sort((a,b)=>b.totalDiscount-a.totalDiscount);}
function detectPromoSameCard(enriched){const byCard={};enriched.forEach(r=>{const card=r._cardNum;if(!card||card==="n/a"||card==="-"||!card.toString().trim())return;if(!r._promoCode)return;if(!byCard[card])byCard[card]={emails:new Set(),rows:[]};if(r._email&&r._email!=="n/a")byCard[card].emails.add(r._email);byCard[card].rows.push(r);});const results=[];Object.entries(byCard).forEach(([card,{emails,rows}])=>{const n=emails.size;if(n<2)return;const totalDiscount=rows.reduce((s,r)=>s+r._discountAmt,0),emailList=[...emails],risk=n>=3?"High":"Mid";const byEmail={};rows.forEach(r=>{if(!r._email)return;if(!byEmail[r._email])byEmail[r._email]={txRows:[],discount:0};byEmail[r._email].txRows.push({orderId:r._orderId,promoCode:r._promoCode,trxFlag:r._trxFlag||"",discountAmt:r._discountAmt,orderStatus:r._orderStatus,createdAt:r._createdAt});byEmail[r._email].discount+=r._discountAmt;});results.push({card,emails:emailList,emailDetails:byEmail,totalDiscount,txCount:rows.length,risk,reasons:[`Card used by ${n} different account${n>1?"s":""} to claim promos: ${emailList.join(", ")}`]});});return results.sort((a,b)=>b.emails.length-a.emails.length||b.totalDiscount-a.totalDiscount);}
function detectPromoSameWallet(enriched){const byWallet={};enriched.forEach(r=>{const wallet=r._walletNum;if(!wallet||wallet==="n/a"||wallet==="-"||!wallet.toString().trim())return;if(!r._promoCode)return;if(!byWallet[wallet])byWallet[wallet]={emails:new Set(),rows:[]};if(r._email&&r._email!=="n/a")byWallet[wallet].emails.add(r._email);byWallet[wallet].rows.push(r);});const results=[];Object.entries(byWallet).forEach(([wallet,{emails,rows}])=>{const n=emails.size;if(n<2)return;const totalDiscount=rows.reduce((s,r)=>s+r._discountAmt,0),emailList=[...emails],risk=n>=3?"High":"Mid";const byEmail={};rows.forEach(r=>{if(!r._email)return;if(!byEmail[r._email])byEmail[r._email]={txRows:[],discount:0};byEmail[r._email].txRows.push({orderId:r._orderId,promoCode:r._promoCode,trxFlag:r._trxFlag||"",discountAmt:r._discountAmt,orderStatus:r._orderStatus,createdAt:r._createdAt});byEmail[r._email].discount+=r._discountAmt;});results.push({wallet,emails:emailList,emailDetails:byEmail,totalDiscount,txCount:rows.length,risk,reasons:[`Wallet used by ${n} different account${n>1?"s":""} to claim promos: ${emailList.join(", ")}`]});});return results.sort((a,b)=>b.emails.length-a.emails.length||b.totalDiscount-a.totalDiscount);}
function detectPromoFakeDomain(enriched){const byEmail={};enriched.forEach(r=>{if(!r._email||!isDisposable(r._email))return;if(!r._promoCode)return;if(!byEmail[r._email])byEmail[r._email]={rows:[],names:new Set(),codes:new Set()};byEmail[r._email].rows.push(r);if(r._name)byEmail[r._email].names.add(r._name);if(r._promoCode)byEmail[r._email].codes.add(r._promoCode);});const results=[];Object.entries(byEmail).forEach(([email,{rows,names,codes}])=>{const totalDiscount=rows.reduce((s,r)=>s+r._discountAmt,0),domain=email.split("@")[1]||"";results.push({email,custNames:[...names],promoCodes:[...codes],domain,totalDiscount,txCount:rows.length,txRows:rows.map(r=>({orderId:r._orderId,promoCode:r._promoCode,trxFlag:r._trxFlag||"",discountAmt:r._discountAmt,orderStatus:r._orderStatus,total:r._total,createdAt:r._createdAt})),reasons:[`Non-whitelisted email domain: @${domain}`]});});return results.sort((a,b)=>b.totalDiscount-a.totalDiscount);}

// ─── File parsing ─────────────────────────────────────────────────────────────
async function parseFile(file){const ext=file.name.split(".").pop().toLowerCase();return new Promise((res,rej)=>{const rdr=new FileReader();if(ext==="csv"){rdr.onload=e=>{const p=Papa.parse(e.target.result,{header:true,skipEmptyLines:true,transformHeader:h=>h.trim()});res(p.data);};rdr.onerror=()=>rej(new Error("Cannot read CSV"));rdr.readAsText(file);}else if(["xlsx","xls"].includes(ext)){rdr.onload=e=>{try{const wb=XLSX.read(e.target.result,{type:"array"});res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""}));}catch(e){rej(e);}};rdr.onerror=()=>rej(new Error("Cannot read Excel"));rdr.readAsArrayBuffer(file);}else rej(new Error("Use .csv, .xlsx or .xls"));});}

// ─── Excel export ─────────────────────────────────────────────────────────────
function makeExcelFile(fraud,highAmt,fakeDom,filename){
  const HEADERS=["Risk","Customer Name","Email","CC / Wallet / Method","Cart/Order ID","Timestamp","Auth Message","Status","Amount (EGP)","Total Amount (EGP)","Reason"];
  const buildRows=(arr,riskLabel)=>arr.flatMap(r=>{const ccs=r.uniqueCCs&&r.uniqueCCs.length>0?r.uniqueCCs:["—"];return ccs.flatMap(cc=>{const txs=(r.ccTxRows&&r.ccTxRows[cc])||[{cartId:"—",timestamp:"—",authMsg:"—",orderStatus:"—",amt:0}];let first=true;return txs.map(tx=>{const row={"Risk":riskLabel||r.risk||"","Customer Name":(r.custNames||[]).join(", ")||"—","Email":r.email||"—","CC / Wallet / Method":cc||"—","Cart/Order ID":tx.cartId||"—","Timestamp":tx.timestamp||"—","Auth Message":tx.authMsg||"—","Status":tx.orderStatus||"—","Amount (EGP)":tx.amt||"","Total Amount (EGP)":first?r.totalAmt:"","Reason":first?(r.reasons||[]).join(" | "):""};first=false;return row;});});});
  const sheets=[{name:"High Risk",rows:buildRows(fraud.filter(r=>r.risk==="High"||r.risk==="HighSuspicious"),"HIGH")},{name:"Mid Risk",rows:buildRows(fraud.filter(r=>r.risk==="Mid"),"MID")},{name:"High Amounts",rows:buildRows(highAmt,"HIGH AMOUNT")},{name:"Fake Domain",rows:buildRows(fakeDom,"FAKE DOMAIN")}];
  const wb=XLSX.utils.book_new();sheets.forEach(({name,rows})=>{const ws=XLSX.utils.json_to_sheet(rows.length>0?rows:[Object.fromEntries(HEADERS.map(h=>[h,""]))],{header:HEADERS});ws["!cols"]=HEADERS.map(()=>({wch:22}));XLSX.utils.book_append_sheet(wb,ws,name);});XLSX.writeFile(wb,`${filename}_${new Date().toISOString().split("T")[0]}.xlsx`);}

// ─── Universal Result Card ────────────────────────────────────────────────────
function ResultCard({r,accentColor,cfg={},activeFilter,customRows,customHeader,onBlock,isBlocked,blockingEmail}){
  const isHigh=r.risk==="High"||r.risk==="HighSuspicious"||r.risk==="HighAmount";
  const isMid=r.risk==="Mid";
  let badgeBg=activeFilter==="fakedomain"?"#581c87":activeFilter==="highamount"?"#065f46":isHigh?"#dc2626":isMid?"#d97706":"#065f46";
  let badgeTxt=activeFilter==="fakedomain"?["FAKE","DOMAIN"]:activeFilter==="highamount"?["HIGH","AMOUNT"]:isHigh?"HIGH":isMid?"MID":"HIGH";

  const ccLabel=cfg.ccLabel||"CC Details";
  const cartLabel=cfg.cartLabel||"Cart ID";
  const showOS=cfg.showOS!==false;
  const showECI=cfg.showECI===true;
  const showCountries=cfg.showCountries===true;
  const hideAuthMsg=cfg.hideAuthMsg===true;
  const hideStatus=cfg.hideStatus===true;
  const showBiller=cfg.showBiller===true;
  const showPayMethod=cfg.showPayMethod===true;

  // Build flat tx rows from ccTxRows
  const allTxRows=[];
  if(r.uniqueCCs&&r.ccTxRows){r.uniqueCCs.forEach(cc=>{const txs=(r.ccTxRows[cc])||[];txs.forEach(tx=>allTxRows.push({cc,...tx}));});}

  const statusClr=s=>{const lw=(s||"").toLowerCase();if(lw==="true"||lw==="captured"||lw==="authorised"||lw==="authorized"||lw==="paid")return{bg:"#dcfce7",cl:"#16a34a"};if(lw==="false"||lw==="declined"||lw==="failed")return{bg:"#fee2e2",cl:"#dc2626"};return{bg:"#f1f5f9",cl:"#64748b"};};
  const authClr=msg=>{if(["Authorised","Authorized"].includes(msg))return"#16a34a";if(msg==="User does not have a wallet")return"#dc2626";return"#475569";};

  const tableRows=customRows||allTxRows;
  const tableHeader=customHeader||[ccLabel,cartLabel,"Timestamp",...(hideAuthMsg?[]:["Auth Message"]),...(hideStatus?[]:["Status"]),...(showBiller?["Service Biller"]:[]),...(showPayMethod?["Payment Method"]:[]),...(showECI?["ECI"]:[]),...(showCountries?["Countries"]:[]),"Amount (EGP)"];

  return(
    <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:`1px solid ${accentColor}22`,marginBottom:16,overflow:"hidden"}}>
      {/* Header strip */}
      <div style={{display:"flex",alignItems:"stretch"}}>
        {/* Risk badge */}
        <div style={{background:badgeBg,width:80,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:2}}>
          {Array.isArray(badgeTxt)?badgeTxt.map((t,i)=><div key={i} style={{color:"#fff",fontWeight:900,fontSize:11,letterSpacing:1,textAlign:"center"}}>{t}</div>):<div style={{color:"#fff",fontWeight:900,fontSize:11,letterSpacing:1}}>{badgeTxt}</div>}
          <div style={{color:"rgba(255,255,255,0.6)",fontSize:10,marginTop:6,fontWeight:600}}>{r.txCount||tableRows.length} tx</div>
        </div>
        {/* Email + names + pills */}
        <div style={{flex:1,padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",borderRight:"1px solid #f1f5f9"}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:5}}>
            {(r.custNames||[]).map((n,i)=><span key={i} style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>{n}</span>)}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{fontSize:13,color:"#334155",fontWeight:500,wordBreak:"break-all"}}>{r.email||r.wallet||""}</span>
            {r.email&&isDisposable(r.email)&&<span style={{background:"#fee2e2",color:"#b91c1c",padding:"2px 7px",borderRadius:4,fontSize:11,fontWeight:700,whiteSpace:"nowrap",flexShrink:0}}>⚠ Non-wl</span>}
          </div>
          {/* CC pills */}
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {(r.uniqueCCs||[]).map((cc,i)=>(
              <div key={i} style={{background:"#1e293b",color:"#a5b4fc",borderRadius:6,padding:"5px 11px",fontSize:12,fontFamily:"'Courier New',monospace",fontWeight:700,display:"inline-flex",alignItems:"center",gap:6}}>
                <span style={{width:6,height:6,borderRadius:"50%",background:accentColor,display:"inline-block",flexShrink:0}}/>
                {cc}
              </div>
            ))}
            {r.wallet&&<div style={{background:"#1e293b",color:"#a5b4fc",borderRadius:6,padding:"5px 11px",fontSize:12,fontFamily:"'Courier New',monospace",fontWeight:700,display:"inline-flex",alignItems:"center",gap:6}}><span style={{width:6,height:6,borderRadius:"50%",background:accentColor,display:"inline-block"}}/>{r.wallet}</div>}
          </div>
        </div>
        {/* Total */}
        <div style={{background:"#f8fafc",padding:"16px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:130,borderRight:"1px solid #f1f5f9"}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Total (EGP)</div>
          <div style={{fontWeight:900,color:accentColor,fontSize:24,lineHeight:1}}>{(r.totalAmt||0).toLocaleString()}</div>
        </div>
        {/* Reason */}
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220,maxWidth:280}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>Reason</div>
          {(r.reasons||[]).map((rs,i)=>(
            <div key={i} style={{display:"flex",gap:6,marginBottom:4}}>
              <span style={{color:accentColor,fontWeight:900,fontSize:14,lineHeight:1,flexShrink:0}}>›</span>
              <span style={{fontSize:12,color:"#475569",lineHeight:1.5}}>{rs}</span>
            </div>
          ))}
        </div>
        {/* Block button */}
        {onBlock&&r.email&&<div style={{padding:"16px 14px",display:"flex",alignItems:"center",justifyContent:"center",borderLeft:"1px solid #f1f5f9"}}>
          <button onClick={()=>onBlock(r.email)} disabled={!!blockingEmail} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4,padding:"10px 14px",background:isBlocked&&isBlocked(r.email)?"#fef2f2":"#fff7ed",border:`1.5px solid ${isBlocked&&isBlocked(r.email)?"#fca5a5":"#fed7aa"}`,borderRadius:10,cursor:"pointer",opacity:blockingEmail===r.email?.toLowerCase()?.trim()?0.5:1,transition:"all 0.15s"}}>
            <span style={{fontSize:16}}>{isBlocked&&isBlocked(r.email)?"🔒":"🚫"}</span>
            <span style={{fontSize:10,fontWeight:700,color:isBlocked&&isBlocked(r.email)?"#dc2626":"#ea580c",whiteSpace:"nowrap"}}>{isBlocked&&isBlocked(r.email)?"Blocked":"Block"}</span>
          </button>
        </div>}
      </div>

      {/* Transactions table */}
      {tableRows.length>0&&(
        <div style={{borderTop:"1px solid #f1f5f9",overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:"#f8fafc"}}>
                {tableHeader.map(h=><th key={h} style={{padding:"8px 14px",textAlign:"left",fontWeight:700,color:"#64748b",fontSize:10,textTransform:"uppercase",letterSpacing:0.6,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {tableRows.map((tx,ti)=>{
                const st=statusClr(tx.orderStatus||tx.status||"");
                const amClr=authClr(tx.authMsg||"");
                return(
                  <tr key={ti} style={{borderBottom:ti<tableRows.length-1?"1px solid #f8fafc":"none",background:ti%2===0?"#fff":"#fafafa"}}>
                    {/* CC pill */}
                    {tx.cc!==undefined&&<td style={{padding:"9px 14px",whiteSpace:"nowrap"}}><div style={{background:"#1e293b",color:"#a5b4fc",borderRadius:5,padding:"3px 9px",fontSize:11,fontFamily:"monospace",fontWeight:700,display:"inline-block"}}>{tx.cc||"—"}</div></td>}
                    {/* Cart ID */}
                    {(tx.cartId!==undefined||tx.txId!==undefined)&&<td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{(tx.cartId||tx.txId)?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{tx.cartId||tx.txId}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td>}
                    {/* Fawry code */}
                    {tx.fawryCode!==undefined&&<td style={{padding:"9px 14px",fontFamily:"monospace",fontWeight:700,color:"#a5b4fc",fontSize:12,whiteSpace:"nowrap"}}>{tx.fawryCode||"—"}</td>}
                    {/* Timestamp */}
                    {tx.timestamp!==undefined&&<td style={{padding:"9px 14px",whiteSpace:"nowrap",fontFamily:"monospace",fontSize:11,color:"#475569",background:"#f8fafc"}}>{tx.timestamp||"—"}</td>}
                    {/* Sold date */}
                    {tx.soldDate!==undefined&&<td style={{padding:"9px 14px",fontSize:11,fontFamily:"monospace",color:"#334155",background:"#f8fafc",whiteSpace:"nowrap"}}>{tx.soldDate||"—"}</td>}
                    {/* Auth message — hidden when null (e.g. Admin tab) */}
                    {tx.authMsg!==undefined&&tx.authMsg!==null&&<td style={{padding:"9px 14px",maxWidth:200}}><span style={{fontSize:12,fontWeight:600,color:amClr}}>{tx.authMsg||"—"}</span></td>}
                    {/* Status */}
                    {!hideStatus&&(tx.orderStatus!==undefined||tx.status!==undefined)&&<td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>
                      <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 9px",borderRadius:6,fontSize:11,fontWeight:700,background:st.bg,color:st.cl}}>
                        <span style={{width:6,height:6,borderRadius:"50%",background:st.cl,display:"inline-block"}}/>
                        {tx.orderStatus||tx.status||"—"}
                      </span>
                    </td>}
                    {/* Service Biller */}
                    {showBiller&&<td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{tx.biller?<span style={{background:"#fef3c7",color:"#92400e",border:"1px solid #fde68a",padding:"2px 9px",borderRadius:6,fontSize:11,fontWeight:600}}>{tx.biller}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td>}
                    {/* Payment Method */}
                    {showPayMethod&&<td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{tx.payMethod?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 9px",borderRadius:6,fontSize:11,fontWeight:600}}>{tx.payMethod}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td>}
                    {/* ECI */}
                    {tx.eci!==undefined&&<td style={{padding:"9px 14px",fontFamily:"monospace",fontWeight:700,color:"#334155",fontSize:12}}>{tx.eci||"—"}</td>}
                    {/* Countries */}
                    {tx.country!==undefined&&<td style={{padding:"9px 14px",fontSize:12,color:"#334155"}}>{tx.country||"—"}</td>}
                    {/* Amount */}
                    <td style={{padding:"9px 14px",textAlign:"right",fontWeight:700,color:"#0f172a",whiteSpace:"nowrap"}}>
                      {(tx.amt||0)>0?<span style={{background:"#f0fdf4",color:"#16a34a",padding:"2px 8px",borderRadius:5,fontSize:12}}>{(tx.amt||0).toLocaleString()}</span>:"—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── Wallet Abuser Card ───────────────────────────────────────────────────────
function WalletAbuserCard({r,accent}){
  const isH=r.risk==="High";
  return(
    <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:`1px solid ${accent}22`,marginBottom:16,overflow:"hidden"}}>
      <div style={{display:"flex",alignItems:"stretch"}}>
        <div style={{background:isH?"#dc2626":"#d97706",width:80,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:2}}>
          <div style={{color:"#fff",fontWeight:900,fontSize:11,letterSpacing:1}}>{r.risk.toUpperCase()}</div>
          <div style={{color:"rgba(255,255,255,0.6)",fontSize:10,marginTop:6,fontWeight:600}}>{r.txCount} tx</div>
        </div>
        <div style={{flex:1,padding:"16px 20px",borderRight:"1px solid #f1f5f9"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <div style={{background:"#1e293b",color:"#a5b4fc",borderRadius:7,padding:"6px 14px",fontSize:13,fontFamily:"monospace",fontWeight:700}}>{r.wallet}</div>
            <span style={{background:"#fef3c7",color:"#92400e",padding:"3px 10px",borderRadius:6,fontSize:11,fontWeight:700}}>👥 {r.emails.length} emails</span>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {r.emails.map((email,ei)=>(
              <span key={ei} style={{fontSize:12,color:"#334155",background:"#f8fafc",borderRadius:5,padding:"3px 9px",border:"1px solid #e2e8f0"}}>
                {email}{isDisposable(email)&&<span style={{marginLeft:5,background:"#fee2e2",color:"#b91c1c",padding:"1px 5px",borderRadius:3,fontSize:10,fontWeight:600}}>⚠</span>}
              </span>
            ))}
          </div>
        </div>
        <div style={{background:"#f8fafc",padding:"16px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:130,borderRight:"1px solid #f1f5f9"}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Total (EGP)</div>
          <div style={{fontWeight:900,color:accent,fontSize:24,lineHeight:1}}>{r.totalAmt.toLocaleString()}</div>
        </div>
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>Reason</div>
          {r.reasons.map((rs,ri)=>(
            <div key={ri} style={{display:"flex",gap:6,marginBottom:4}}><span style={{color:accent,fontWeight:900,fontSize:14}}>›</span><span style={{fontSize:12,color:"#475569",lineHeight:1.5}}>{rs}</span></div>
          ))}
        </div>
      </div>
      {/* Per-email breakdown */}
      {r.emails.map((email,ei)=>{
        const d=r.emailDetails[email]||{};
        return(
          <div key={ei} style={{borderTop:"1px solid #f1f5f9"}}>
            <div style={{padding:"8px 16px",background:"#f8fafc",display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:11,fontWeight:700,color:"#475569"}}>📧</span>
              <span style={{fontSize:12,color:"#334155",fontWeight:600}}>{email}</span>
              <span style={{fontSize:11,color:"#94a3b8"}}>· {(d.txRows||[]).length} transactions · EGP {(d.amt||0).toLocaleString()}</span>
            </div>
            {(d.txRows||[]).length>0&&(
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#fff"}}>{["Order ID","Timestamp","Auth Message","Status","Amount (EGP)"].map(h=><th key={h} style={{padding:"6px 14px",textAlign:"left",fontWeight:700,color:"#94a3b8",fontSize:10,textTransform:"uppercase",letterSpacing:0.5,borderBottom:"1px solid #f1f5f9",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>{(d.txRows||[]).map((tx,ti)=>{const st=(tx.orderStatus||"").toLowerCase();const isS=st==="true";const isF=st==="false";return(<tr key={ti} style={{borderBottom:ti<(d.txRows||[]).length-1?"1px solid #f8fafc":"none",background:ti%2===0?"#fff":"#fafafa"}}>
                  <td style={{padding:"8px 14px",whiteSpace:"nowrap"}}>{tx.cartId?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{tx.cartId}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td>
                  <td style={{padding:"8px 14px",fontFamily:"monospace",fontSize:11,color:"#475569",background:"#f8fafc",whiteSpace:"nowrap"}}>{tx.timestamp||"—"}</td>
                  <td style={{padding:"8px 14px",fontSize:12,fontWeight:600,color:["Authorised","Authorized"].includes(tx.authMsg)?"#16a34a":tx.authMsg==="User does not have a wallet"?"#dc2626":"#475569"}}>{tx.authMsg||"—"}</td>
                  <td style={{padding:"8px 14px",whiteSpace:"nowrap"}}><span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 9px",borderRadius:6,fontSize:11,fontWeight:700,background:isS?"#dcfce7":isF?"#fee2e2":"#f1f5f9",color:isS?"#16a34a":isF?"#dc2626":"#64748b"}}><span style={{width:6,height:6,borderRadius:"50%",background:isS?"#16a34a":isF?"#dc2626":"#94a3b8",display:"inline-block"}}/>{tx.orderStatus||"—"}</span></td>
                  <td style={{padding:"8px 14px",textAlign:"right",fontWeight:700}}>{tx.amt>0?<span style={{background:"#f0fdf4",color:"#16a34a",padding:"2px 8px",borderRadius:5,fontSize:12}}>{tx.amt.toLocaleString()}</span>:"—"}</td>
                </tr>);})}
                </tbody>
              </table>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Universal Dashboard ──────────────────────────────────────────────────────
function UniversalDashboard({config}){
  const{accent,reimportLabel,raw,enriched,tabs,stats,onReimport,showRaw,setShowRaw,previewCols,previewKeyFn,onDownload,onBlock,isBlocked,blockingEmail}=config;
  const [activeTab,setActiveTab]=useState(tabs[0].id);
  const [search,setSearch]=useState("");

  const activeTabCfg=tabs.find(t=>t.id===activeTab)||tabs[0];
  const baseRows=activeTabCfg.rows||[];
  const filteredRows=search?baseRows.filter(r=>{const q=search.toLowerCase();return(r.email||"").includes(q)||(r.custNames||[]).some(n=>n.toLowerCase().includes(q))||(r.uniqueCCs||[]).some(c=>c.toLowerCase().includes(q))||(r.wallet||"").toLowerCase().includes(q)||(r.userId||"").toLowerCase().includes(q)||(r.recharge||"").toLowerCase().includes(q);}):baseRows;
  const activeAccent=activeTabCfg.accent||accent;
  const uniqueEmails=new Set(enriched.map(r=>r._email||"").filter(Boolean)).size;

  return(<>
    {/* Header */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
      <div><div style={{fontWeight:800,fontSize:18,color:"#0f172a"}}>{reimportLabel} Fraud Report</div><div style={{fontSize:13,color:"#64748b",marginTop:3}}>{raw.length.toLocaleString()} records · {uniqueEmails.toLocaleString()} unique customers</div></div>
      <button onClick={onReimport} style={{display:"flex",alignItems:"center",gap:7,padding:"9px 18px",background:accent,color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}><Upload size={13}/> Re-import</button>
    </div>

    {/* Stats */}
    <div style={{display:"grid",gridTemplateColumns:`repeat(${stats.length},1fr)`,gap:14,marginBottom:24}}>
      {stats.map(({label,value,clr,Icon})=>(
        <div key={label} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderLeft:`5px solid ${clr}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
            <span style={{fontSize:11,color:"#64748b",fontWeight:600,textTransform:"uppercase",letterSpacing:0.5}}>{label}</span>
            <div style={{width:30,height:30,borderRadius:7,background:clr+"18",display:"flex",alignItems:"center",justifyContent:"center"}}><Icon size={14} color={clr}/></div>
          </div>
          <div style={{fontSize:30,fontWeight:900,color:clr,lineHeight:1}}>{value}</div>
        </div>
      ))}
    </div>

    {/* Filter + Search bar */}
    <div style={{background:"#fff",borderRadius:14,padding:"14px 18px",marginBottom:16,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
      <div style={{display:"flex",gap:4,flexWrap:"wrap",flex:1}}>
        {tabs.map(t=>(
          <button key={t.id} onClick={()=>{setActiveTab(t.id);setSearch("");}} style={{padding:"6px 14px",borderRadius:8,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:activeTab===t.id?(t.accent||accent):"#f1f5f9",color:activeTab===t.id?"#fff":"#64748b",display:"flex",alignItems:"center",gap:5}}>
            {t.label}<span style={{background:activeTab===t.id?"rgba(255,255,255,0.25)":"#e2e8f0",color:activeTab===t.id?"#fff":"#64748b",fontSize:10,padding:"1px 6px",borderRadius:10,fontWeight:700}}>{t.rows.length}</span>
          </button>
        ))}
      </div>
      <div style={{position:"relative"}}>
        <Search size={12} style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",color:"#94a3b8"}}/>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search…" style={{paddingLeft:27,paddingRight:10,paddingTop:7,paddingBottom:7,borderRadius:8,border:"1px solid #e2e8f0",fontSize:12,outline:"none",width:200,color:"#334155"}}/>
      </div>
      {onDownload&&<button onClick={onDownload} disabled={filteredRows.length===0} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 14px",background:filteredRows.length===0?"#e2e8f0":"#16a34a",color:filteredRows.length===0?"#94a3b8":"#fff",border:"none",borderRadius:7,cursor:filteredRows.length===0?"not-allowed":"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}><Download size={13}/> Download .xlsx</button>}
    </div>

    {/* Cards */}
    {filteredRows.length===0?(
      <div style={{background:"#fff",borderRadius:16,padding:60,textAlign:"center",color:"#94a3b8",fontSize:14,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>✅ No results in this category.</div>
    ):activeTabCfg.renderCard?(
      <div>{filteredRows.map((r,i)=>activeTabCfg.renderCard(r,i,activeAccent))}</div>
    ):(
      <div>{filteredRows.map((r,i)=><ResultCard key={i} r={r} accentColor={activeAccent} cfg={activeTabCfg.cfg||{}} activeFilter={activeTab} onBlock={onBlock} isBlocked={isBlocked} blockingEmail={blockingEmail}/>)}</div>
    )}

    {/* Raw data */}
    <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",overflow:"hidden",marginTop:16}}>
      <button onClick={()=>setShowRaw(v=>!v)} style={{width:"100%",padding:"15px 22px",display:"flex",alignItems:"center",gap:10,background:"none",border:"none",cursor:"pointer",textAlign:"left"}}>
        <Eye size={15} color={accent}/><span style={{fontWeight:700,color:"#0f172a",fontSize:14}}>Full Transaction Data</span><span style={{fontSize:12,color:"#94a3b8"}}>({raw.length} records)</span>
        <span style={{marginLeft:"auto",fontSize:12,color:accent,fontWeight:600}}>{showRaw?"▲ Hide":"▼ Show"}</span>
      </button>
      {showRaw&&(
        <div style={{overflowX:"auto",maxHeight:380,overflowY:"auto",borderTop:"1px solid #f1f5f9"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:"#1e293b",position:"sticky",top:0,zIndex:1}}>
              {previewCols.map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",fontWeight:600,color:"#94a3b8",whiteSpace:"nowrap",fontSize:11,textTransform:"uppercase",letterSpacing:0.5}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {enriched.slice(0,300).map((row,i)=>(
                <tr key={i} style={{borderBottom:"1px solid #f8fafc",background:i%2===0?"#fff":"#fafafa"}}>
                  {previewCols.map(col=><td key={col} style={{padding:"8px 14px",color:"#334155",whiteSpace:"nowrap",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",fontSize:13}}>{previewKeyFn?previewKeyFn(row,col):getCol(row)(col)||"—"}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
          {raw.length>300&&<div style={{padding:"9px 16px",fontSize:12,color:"#94a3b8",borderTop:"1px solid #f1f5f9"}}>Showing 300 of {raw.length} rows</div>}
        </div>
      )}
    </div>
  </>);
}

// ─── Import Modal ─────────────────────────────────────────────────────────────
function ImportModal({open,title,accent,onClose,onImport,rows,setRows,step,setStep,previewCols,onFileDrop,loading,setLoading,err,setErr}){
  const fRef=useRef();
  const [drag,setDrag]=useState(false);
  const doFile=useCallback(async file=>{setLoading(true);setErr("");try{const r=await parseFile(file);setRows(r);setStep("preview");}catch(e){setErr(e.message);}finally{setLoading(false);};},[setLoading,setErr,setRows,setStep]);
  if(!open)return null;
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
    <div style={{background:"#fff",borderRadius:18,width:520,maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 30px 90px rgba(0,0,0,0.35)",overflow:"hidden"}}>
      <div style={{padding:"18px 24px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center",background:"#fafafa"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}><div style={{width:10,height:10,borderRadius:"50%",background:accent}}/><span style={{fontWeight:700,fontSize:16,color:"#0f172a"}}>{title}</span></div>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8"}}><X size={20}/></button>
      </div>
      <div style={{flex:1,overflow:"auto",padding:24}}>
        {step==="drop"&&(<>
          <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);if(e.dataTransfer.files[0])doFile(e.dataTransfer.files[0]);}} onClick={()=>fRef.current.click()} style={{border:`2px dashed ${drag?accent:"#e2e8f0"}`,borderRadius:14,padding:52,textAlign:"center",cursor:"pointer",background:drag?accent+"11":"#fafafa"}}>
            {loading?<div style={{color:"#64748b",fontSize:14}}>Reading file…</div>:<>
              <div style={{width:56,height:56,borderRadius:14,background:drag?accent+"22":"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px"}}><Upload size={24} color={drag?accent:"#94a3b8"}/></div>
              <div style={{fontWeight:700,color:"#1e293b",fontSize:15,marginBottom:6}}>Drop your file here</div>
              <div style={{fontSize:13,color:"#94a3b8"}}>CSV · Excel (.xlsx / .xls)</div>
            </>}
            <input ref={fRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])doFile(e.target.files[0]);e.target.value="";}}/>
          </div>
          {err&&<div style={{marginTop:14,background:"#fef2f2",borderRadius:10,padding:14,color:"#dc2626",fontSize:13}}>⚠️ {err}</div>}
        </>)}
        {step==="preview"&&(<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:18}}>
            {[["Rows",rows.length,"#0f172a"],["Unique Emails",new Set(rows.map(r=>getCol(r)("client_email","Customer Email","User Email","user email","email","user_email").toLowerCase()).filter(Boolean)).size,accent],["Columns",Object.keys(rows[0]||{}).length,"#64748b"]].map(([l,v,c])=>(
              <div key={l} style={{background:"#f8fafc",borderRadius:10,padding:"12px",textAlign:"center"}}><div style={{fontSize:24,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:12,color:"#64748b",marginTop:3}}>{l}</div></div>
            ))}
          </div>
          <div style={{border:"1px solid #e2e8f0",borderRadius:10,overflow:"hidden",maxHeight:220,overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead><tr style={{background:"#1e293b",position:"sticky",top:0}}>{previewCols.slice(0,6).map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontWeight:600,color:"#94a3b8",whiteSpace:"nowrap",fontSize:11}}>{h}</th>)}</tr></thead>
              <tbody>{rows.slice(0,8).map((r,i)=>(<tr key={i} style={{borderBottom:"1px solid #f8fafc"}}>{previewCols.slice(0,6).map(col=><td key={col} style={{padding:"6px 10px",color:"#334155",whiteSpace:"nowrap",maxWidth:130,overflow:"hidden",textOverflow:"ellipsis"}}>{getCol(r)(col)||"—"}</td>)}</tr>))}</tbody>
            </table>
          </div>
        </>)}
        {step==="done"&&(<div style={{textAlign:"center",padding:"30px 0"}}>
          <div style={{width:64,height:64,borderRadius:"50%",background:"#dcfce7",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}><CheckCircle size={30} color="#16a34a"/></div>
          <div style={{fontWeight:800,fontSize:18,color:"#0f172a",marginBottom:10}}>Analysis Complete!</div>
        </div>)}
      </div>
      <div style={{padding:"14px 24px",borderTop:"1px solid #f1f5f9",display:"flex",gap:10,justifyContent:"flex-end",background:"#fafafa"}}>
        {step==="drop"&&<button onClick={onClose} style={{padding:"9px 20px",borderRadius:9,border:"1px solid #e2e8f0",background:"#fff",cursor:"pointer",fontSize:13,color:"#64748b"}}>Cancel</button>}
        {step==="preview"&&<><button onClick={()=>setStep("drop")} style={{padding:"9px 20px",borderRadius:9,border:"1px solid #e2e8f0",background:"#fff",cursor:"pointer",fontSize:13,color:"#64748b"}}>← Back</button><button onClick={onImport} style={{padding:"9px 26px",borderRadius:9,border:"none",background:accent,color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>Run Fraud Analysis →</button></>}
        {step==="done"&&<button onClick={onClose} style={{padding:"9px 24px",borderRadius:9,border:"none",background:"#16a34a",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>View Results</button>}
      </div>
    </div>
  </div>);
}

// ─── Auth screens ─────────────────────────────────────────────────────────────
function ForgotPasswordScreen({onBack}){
  const [step,setStep]=useState("username");const [username,setUsername]=useState("");const [answer,setAnswer]=useState("");const [newPw,setNewPw]=useState("");const [confirmPw,setConfirmPw]=useState("");const [showPw,setShowPw]=useState(false);const [error,setError]=useState("");const [foundUser,setFoundUser]=useState(null);
  const checkUsername=()=>{setError("");const users=loadUsers();const u=users.find(u=>u.username===username.trim());if(!u){setError("Username not found.");return;}if(!u.securityAnswer){setError("No security question set. Contact your admin.");return;}setFoundUser(u);setStep("question");};
  const checkAnswer=()=>{setError("");if(answer.trim().toLowerCase()!==foundUser.securityAnswer){setError("Incorrect answer.");return;}setStep("newpw");};
  const resetPassword=()=>{setError("");if(newPw.length<6){setError("Password must be at least 6 characters.");return;}if(newPw!==confirmPw){setError("Passwords don't match.");return;}const users=loadUsers();saveUsers(users.map(u=>u.username===foundUser.username?{...u,password:newPw}:u));setStep("done");};
  return(<div style={{minHeight:"100vh",background:"#0f172a",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',sans-serif"}}>
    <div style={{marginBottom:28,textAlign:"center"}}><div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginBottom:8}}><Shield size={28} color="#f97316"/><span style={{fontWeight:900,fontSize:22,color:"#f8fafc"}}>Waffarha Guard</span></div></div>
    <div style={{background:"#1e293b",borderRadius:20,padding:"36px 40px",width:400,boxShadow:"0 30px 80px rgba(0,0,0,0.5)"}}>
      {step==="username"&&<><div style={{fontWeight:700,fontSize:17,color:"#f8fafc",marginBottom:20}}>Forgot your password?</div><div style={{marginBottom:16}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>Username</label><input value={username} onChange={e=>setUsername(e.target.value)} onKeyDown={e=>e.key==="Enter"&&checkUsername()} placeholder="Enter username" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>{error&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {error}</div>}<button onClick={checkUsername} disabled={!username} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:!username?"#334155":"#f97316",color:!username?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:!username?"not-allowed":"pointer",marginBottom:12}}>Next →</button><button onClick={onBack} style={{width:"100%",padding:"11px",borderRadius:10,border:"1px solid #334155",background:"none",color:"#64748b",fontWeight:600,fontSize:13,cursor:"pointer"}}>← Back to Login</button></>}
      {step==="question"&&<><div style={{fontWeight:700,fontSize:17,color:"#f8fafc",marginBottom:6}}>Security Question</div><div style={{background:"#0f172a",borderRadius:10,padding:"12px 16px",marginBottom:18,border:"1px solid #334155"}}><div style={{fontSize:14,color:"#f97316",fontWeight:700}}>{SECURITY_QUESTION}</div></div><div style={{marginBottom:16}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>Your Answer</label><input value={answer} onChange={e=>setAnswer(e.target.value)} onKeyDown={e=>e.key==="Enter"&&checkAnswer()} placeholder="Enter your answer" autoComplete="off" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>{error&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {error}</div>}<button onClick={checkAnswer} disabled={!answer} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:!answer?"#334155":"#f97316",color:!answer?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:!answer?"not-allowed":"pointer",marginBottom:12}}>Verify →</button><button onClick={()=>{setStep("username");setError("");}} style={{width:"100%",padding:"11px",borderRadius:10,border:"1px solid #334155",background:"none",color:"#64748b",fontWeight:600,fontSize:13,cursor:"pointer"}}>← Back</button></>}
      {step==="newpw"&&<><div style={{fontWeight:700,fontSize:17,color:"#f8fafc",marginBottom:18}}>Set New Password</div>{[["New Password",newPw,setNewPw],["Confirm New Password",confirmPw,setConfirmPw]].map(([lbl,val,setter])=>(<div key={lbl} style={{marginBottom:14}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>{lbl}</label><input value={val} onChange={e=>setter(e.target.value)} type="password" placeholder="••••••••" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>))}{error&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {error}</div>}<button onClick={resetPassword} disabled={!newPw||!confirmPw} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:(!newPw||!confirmPw)?"#334155":"#16a34a",color:(!newPw||!confirmPw)?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:(!newPw||!confirmPw)?"not-allowed":"pointer"}}>Reset Password</button></>}
      {step==="done"&&<div style={{textAlign:"center",padding:"10px 0"}}><div style={{fontSize:48,marginBottom:16}}>✅</div><div style={{fontWeight:800,fontSize:18,color:"#f8fafc",marginBottom:10}}>Password Reset!</div><button onClick={onBack} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:"#f97316",color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer"}}>← Back to Login</button></div>}
    </div>
  </div>);
}

function LoginScreen({onLogin}){
  const [username,setUsername]=useState("");const [password,setPassword]=useState("");const [showPw,setShowPw]=useState(false);const [error,setError]=useState("");const [loading,setLoading]=useState(false);const [screen,setScreen]=useState("login");
  if(screen==="forgot")return <ForgotPasswordScreen onBack={()=>setScreen("login")}/>;
  const handleLogin=()=>{setError("");setLoading(true);setTimeout(()=>{const u=username.trim();if(u===SUPER_ADMIN.username&&password===getSuperAdminPassword()){onLogin({username:u,role:"superadmin"});return;}const users=loadUsers();const found=users.find(u2=>u2.username===u&&u2.password===password);if(found){onLogin({username:u,role:"user",hasSecurityAnswer:!!found.securityAnswer});return;}setError("Invalid username or password.");setLoading(false);},600);};
  return(<div style={{minHeight:"100vh",background:"#0f172a",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',sans-serif"}}>
    <div style={{marginBottom:32,textAlign:"center"}}><div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginBottom:4}}><Shield size={32} color="#f97316"/><span style={{fontWeight:900,fontSize:24,color:"#f8fafc"}}>Waffarha Guard</span></div><div style={{fontWeight:800,fontSize:13,color:"#f97316",letterSpacing:0.5,marginBottom:8}}>Waffarha</div><div style={{fontSize:12,color:"#475569",background:"#1e293b",padding:"3px 12px",borderRadius:20,display:"inline-block"}}>Fraud Control</div></div>
    <div style={{background:"#1e293b",borderRadius:20,padding:"36px 40px",width:360,boxShadow:"0 30px 80px rgba(0,0,0,0.5)"}}>
      <div style={{fontWeight:700,fontSize:18,color:"#f8fafc",marginBottom:6}}>Sign in</div><div style={{fontSize:13,color:"#64748b",marginBottom:28}}>Enter your credentials to access the platform</div>
      <div style={{marginBottom:16}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>Username</label><input value={username} onChange={e=>setUsername(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="Enter username" autoComplete="username" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>
      <div style={{marginBottom:10}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>Password</label><div style={{position:"relative"}}><input value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} type={showPw?"text":"password"} placeholder="••••••••" autoComplete="current-password" style={{width:"100%",padding:"11px 40px 11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/><button onClick={()=>setShowPw(v=>!v)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#475569",fontSize:13}}>{showPw?"🙈":"👁"}</button></div></div>
      <div style={{textAlign:"right",marginBottom:20}}><button onClick={()=>setScreen("forgot")} style={{background:"none",border:"none",cursor:"pointer",color:"#f97316",fontSize:12,fontWeight:600,padding:0}}>Forgot password?</button></div>
      {error&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {error}</div>}
      <button onClick={handleLogin} disabled={loading||!username||!password} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:(!username||!password)?"#334155":"#f97316",color:(!username||!password)?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:(!username||!password)?"not-allowed":"pointer",boxShadow:username&&password?"0 4px 14px rgba(249,115,22,0.4)":"none"}}>{loading?"Signing in…":"Sign in →"}</button>
    </div>
  </div>);
}

function AdminPanel({onClose}){
  const [users,setUsers]=useState(loadUsers());const [newUser,setNewUser]=useState({username:"",password:"",role:"user"});const [editIdx,setEditIdx]=useState(null);const [editData,setEditData]=useState({});const [showPw,setShowPw]=useState({});const [msg,setMsg]=useState("");const [activeTab,setActiveTab]=useState("users");const [myPw,setMyPw]=useState({current:"",newPw:"",confirm:""});const [myPwErr,setMyPwErr]=useState("");
  const flash=m=>{setMsg(m);setTimeout(()=>setMsg(""),2500);};const persist=u=>{saveUsers(u);setUsers(u);};
  const addUser=()=>{if(!newUser.username.trim()||!newUser.password.trim())return;if(users.find(u=>u.username===newUser.username.trim())||newUser.username.trim()===SUPER_ADMIN.username){flash("Username already exists.");return;}persist([...users,{username:newUser.username.trim(),password:newUser.password.trim(),role:newUser.role}]);setNewUser({username:"",password:"",role:"user"});flash("User added.");};
  const deleteUser=i=>{persist(users.filter((_,idx)=>idx!==i));flash("User deleted.");};
  const startEdit=i=>{setEditIdx(i);setEditData({...users[i]});};
  const saveEdit=()=>{if(!editData.username.trim()||!editData.password.trim())return;persist(users.map((u,i)=>i===editIdx?{...u,username:editData.username.trim(),password:editData.password.trim(),role:editData.role}:u));setEditIdx(null);flash("User updated.");};
  const updateMyPassword=()=>{setMyPwErr("");if(myPw.current!==getSuperAdminPassword()){setMyPwErr("Current password is incorrect.");return;}if(myPw.newPw.length<6){setMyPwErr("New password must be at least 6 characters.");return;}if(myPw.newPw!==myPw.confirm){setMyPwErr("New passwords don't match.");return;}setSuperAdminPassword(myPw.newPw);setMyPw({current:"",newPw:"",confirm:""});flash("Password updated.");setActiveTab("users");};
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,fontFamily:"'Inter',sans-serif"}}>
    <div style={{background:"#1e293b",borderRadius:20,width:640,maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 30px 90px rgba(0,0,0,0.6)",overflow:"hidden"}}>
      <div style={{padding:"18px 24px",borderBottom:"1px solid #334155",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:10}}><Users size={18} color="#f97316"/><span style={{fontWeight:700,fontSize:16,color:"#f8fafc"}}>User Management</span><span style={{fontSize:11,background:"#f97316",color:"#fff",padding:"2px 8px",borderRadius:20,fontWeight:700}}>SUPER ADMIN</span></div><button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#64748b"}}><X size={20}/></button></div>
      <div style={{display:"flex",gap:0,padding:"12px 24px 0",borderBottom:"1px solid #334155"}}>{[["users","👥 Users"],["mypassword","🔑 My Password"]].map(([v,l])=>(<button key={v} onClick={()=>setActiveTab(v)} style={{padding:"8px 20px",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:"transparent",color:activeTab===v?"#f97316":"#64748b",borderBottom:activeTab===v?"2px solid #f97316":"2px solid transparent",marginBottom:-1}}>{l}</button>))}</div>
      <div style={{flex:1,overflow:"auto",padding:24}}>
        {msg&&<div style={{background:"#dcfce7",border:"1px solid #86efac",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#16a34a",marginBottom:16}}>✓ {msg}</div>}
        {activeTab==="users"&&<>
          <div style={{background:"#0f172a",borderRadius:14,padding:"18px 20px",marginBottom:24,border:"1px solid #334155"}}>
            <div style={{fontWeight:700,color:"#94a3b8",fontSize:12,textTransform:"uppercase",letterSpacing:0.5,marginBottom:14}}>Add New User</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 140px auto",gap:10,alignItems:"flex-end"}}>
              {[["Username",newUser.username,v=>setNewUser(u=>({...u,username:v}))],["Password",newUser.password,v=>setNewUser(u=>({...u,password:v}))]].map(([lbl,val,setter])=>(<div key={lbl}><label style={{fontSize:11,color:"#64748b",fontWeight:600,display:"block",marginBottom:4}}>{lbl}</label><input value={val} onChange={e=>setter(e.target.value)} placeholder={lbl.toLowerCase()} style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #334155",background:"#1e293b",color:"#f8fafc",fontSize:13,outline:"none",boxSizing:"border-box"}}/></div>))}
              <div><label style={{fontSize:11,color:"#64748b",fontWeight:600,display:"block",marginBottom:4}}>Role</label><select value={newUser.role} onChange={e=>setNewUser(v=>({...v,role:e.target.value}))} style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #334155",background:"#1e293b",color:"#f8fafc",fontSize:13,outline:"none",boxSizing:"border-box"}}><option value="user">User</option><option value="superadmin">Super Admin</option></select></div>
              <button onClick={addUser} disabled={!newUser.username||!newUser.password} style={{padding:"9px 18px",borderRadius:8,border:"none",background:(!newUser.username||!newUser.password)?"#334155":"#f97316",color:(!newUser.username||!newUser.password)?"#64748b":"#fff",fontWeight:700,fontSize:13,cursor:(!newUser.username||!newUser.password)?"not-allowed":"pointer"}}>+ Add</button>
            </div>
          </div>
          <div style={{fontWeight:700,color:"#94a3b8",fontSize:12,textTransform:"uppercase",letterSpacing:0.5,marginBottom:12}}>Users ({users.length})</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            <div style={{background:"#0f172a",borderRadius:12,padding:"14px 16px",border:"1px solid #f97316",display:"flex",alignItems:"center",gap:12}}><div style={{flex:1}}><div style={{fontWeight:700,color:"#f8fafc",fontSize:14}}>{SUPER_ADMIN.username}</div><div style={{fontSize:11,color:"#64748b",marginTop:2}}>••••••••••••</div></div><span style={{fontSize:11,background:"#f97316",color:"#fff",padding:"2px 8px",borderRadius:20,fontWeight:700}}>SUPER ADMIN</span></div>
            {users.map((u,i)=>(<div key={i} style={{background:"#0f172a",borderRadius:12,padding:"14px 16px",border:"1px solid #334155"}}>
              {editIdx===i?(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 130px auto auto",gap:8,alignItems:"center"}}><input value={editData.username} onChange={e=>setEditData(v=>({...v,username:e.target.value}))} style={{padding:"7px 10px",borderRadius:7,border:"1px solid #334155",background:"#1e293b",color:"#f8fafc",fontSize:13,outline:"none"}}/><input value={editData.password} onChange={e=>setEditData(v=>({...v,password:e.target.value}))} type="text" style={{padding:"7px 10px",borderRadius:7,border:"1px solid #334155",background:"#1e293b",color:"#f8fafc",fontSize:13,outline:"none"}}/><select value={editData.role||"user"} onChange={e=>setEditData(v=>({...v,role:e.target.value}))} style={{padding:"7px 10px",borderRadius:7,border:"1px solid #334155",background:"#1e293b",color:"#f8fafc",fontSize:13,outline:"none"}}><option value="user">User</option><option value="superadmin">Super Admin</option></select><button onClick={saveEdit} style={{padding:"7px 14px",borderRadius:7,border:"none",background:"#16a34a",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>Save</button><button onClick={()=>setEditIdx(null)} style={{padding:"7px 14px",borderRadius:7,border:"1px solid #334155",background:"none",color:"#94a3b8",fontWeight:700,fontSize:12,cursor:"pointer"}}>Cancel</button></div>)
              :(<div style={{display:"flex",alignItems:"center",gap:12}}><div style={{flex:1}}><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:700,color:"#f8fafc",fontSize:14}}>{u.username}</span>{u.role==="superadmin"&&<span style={{fontSize:10,background:"#f97316",color:"#fff",padding:"1px 7px",borderRadius:20,fontWeight:700}}>SUPER ADMIN</span>}</div><div style={{fontSize:11,color:"#64748b",marginTop:2,fontFamily:"monospace"}}>{showPw[i]?u.password:"••••••••••••"}</div></div><button onClick={()=>setShowPw(v=>({...v,[i]:!v[i]}))} style={{background:"none",border:"none",cursor:"pointer",color:"#475569",fontSize:13,padding:"4px"}}>{showPw[i]?"🙈":"👁"}</button><button onClick={()=>startEdit(i)} style={{padding:"6px 14px",borderRadius:7,border:"1px solid #334155",background:"none",color:"#94a3b8",fontWeight:600,fontSize:12,cursor:"pointer"}}>Edit</button><button onClick={()=>deleteUser(i)} style={{padding:"6px 14px",borderRadius:7,border:"none",background:"#dc2626",color:"#fff",fontWeight:600,fontSize:12,cursor:"pointer"}}>Delete</button></div>)}
            </div>))}
          </div>
        </>}
        {activeTab==="mypassword"&&<div style={{maxWidth:380}}>
          <div style={{fontWeight:700,color:"#f8fafc",fontSize:15,marginBottom:18}}>Change Your Password</div>
          {[["Current Password",myPw.current,v=>setMyPw(p=>({...p,current:v}))],["New Password",myPw.newPw,v=>setMyPw(p=>({...p,newPw:v}))],["Confirm New Password",myPw.confirm,v=>setMyPw(p=>({...p,confirm:v}))]].map(([lbl,val,setter])=>(<div key={lbl} style={{marginBottom:14}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>{lbl}</label><input value={val} onChange={e=>setter(e.target.value)} type="password" placeholder="••••••••" autoComplete="off" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #475569",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>))}
          {myPwErr&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {myPwErr}</div>}
          <button onClick={updateMyPassword} disabled={!myPw.current||!myPw.newPw||!myPw.confirm} style={{padding:"11px 28px",borderRadius:10,border:"none",background:(!myPw.current||!myPw.newPw||!myPw.confirm)?"#334155":"#f97316",color:(!myPw.current||!myPw.newPw||!myPw.confirm)?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:(!myPw.current||!myPw.newPw||!myPw.confirm)?"not-allowed":"pointer"}}>Update Password</button>
        </div>}
      </div>
    </div>
  </div>);
}

function ChangeMyPasswordModal({currentUser,onClose}){
  const [cur,setCur]=useState("");const [newPw,setNewPw]=useState("");const [confirm,setConfirm]=useState("");const [showPw,setShowPw]=useState(false);const [error,setError]=useState("");const [done,setDone]=useState(false);
  const handleSubmit=()=>{setError("");const users=loadUsers();const u=users.find(u=>u.username===currentUser.username);if(!u||u.password!==cur){setError("Current password is incorrect.");return;}if(newPw.length<6){setError("New password must be at least 6 characters.");return;}if(newPw!==confirm){setError("Passwords don't match.");return;}saveUsers(users.map(u=>u.username===currentUser.username?{...u,password:newPw}:u));setDone(true);};
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:2000,fontFamily:"'Inter',sans-serif"}}>
    <div style={{background:"#1e293b",borderRadius:20,width:400,boxShadow:"0 30px 90px rgba(0,0,0,0.6)",overflow:"hidden"}}>
      <div style={{padding:"18px 24px",borderBottom:"1px solid #334155",display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:700,fontSize:16,color:"#f8fafc"}}>🔑 Change Password</span><button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#64748b"}}><X size={20}/></button></div>
      <div style={{padding:24}}>{done?(<div style={{textAlign:"center",padding:"10px 0"}}><div style={{fontSize:40,marginBottom:12}}>✅</div><div style={{fontWeight:800,fontSize:16,color:"#f8fafc",marginBottom:8}}>Password Updated!</div><button onClick={onClose} style={{padding:"10px 28px",borderRadius:10,border:"none",background:"#f97316",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>Done</button></div>):(<>{[["Current Password",cur,setCur],["New Password",newPw,setNewPw],["Confirm New Password",confirm,setConfirm]].map(([label,val,setter])=>(<div key={label} style={{marginBottom:14}}><label style={{fontSize:12,fontWeight:600,color:"#94a3b8",display:"block",marginBottom:6,textTransform:"uppercase"}}>{label}</label><input value={val} onChange={e=>setter(e.target.value)} type={showPw?"text":"password"} placeholder="••••••••" style={{width:"100%",padding:"11px 14px",borderRadius:10,border:"1px solid #334155",background:"#0f172a",color:"#f8fafc",fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>))}{error&&<div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:8,padding:"9px 14px",fontSize:13,color:"#dc2626",marginBottom:16}}>⚠ {error}</div>}<button onClick={handleSubmit} disabled={!cur||!newPw||!confirm} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",background:(!cur||!newPw||!confirm)?"#334155":"#f97316",color:(!cur||!newPw||!confirm)?"#64748b":"#fff",fontWeight:700,fontSize:14,cursor:(!cur||!newPw||!confirm)?"not-allowed":"pointer"}}>Update Password</button></>)}</div>
    </div>
  </div>);
}

// ─── Sidebar tabs ─────────────────────────────────────────────────────────────
const PLATFORM_TABS=[
  {id:"paytabs",label:"PayTabs",clr:"#7c3aed",ready:true},
  {id:"noon",label:"Noon",clr:"#f59e0b",ready:true},
  {id:"paymob",label:"PayMob",clr:"#2563eb",ready:true},
  {id:"admin",label:"Admin",clr:"#0ea5e9",ready:true},
  {id:"fawry",label:"Fawry",clr:"#f97316",ready:true},
  {id:"promo",label:"Promo Abusers",clr:"#e11d48",ready:true},
  {id:"dashboard",label:"📊 Dashboard",clr:"#6366f1",ready:true,adminOnly:true},
  {id:"audit",label:"Audit Log",clr:"#10b981",ready:true,adminOnly:true},
  {id:"settings",label:"⚙ Settings",clr:"#64748b",ready:true,adminOnly:true},
];

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App(){
  const [currentUser,setCurrentUser]=useState(null);
  const [showAdminPanel,setShowAdminPanel]=useState(false);
  const [showChangePw,setShowChangePw]=useState(false);
  const [sidebarOpen,setSidebarOpen]=useState(true);
  const [tab,setTab]=useState("paytabs");

  // ── PayTabs state ──
  const [ptRaw,setPtRaw]=useState([]);const [ptEnriched,setPtEnriched]=useState([]);const [ptFraud,setPtFraud]=useState([]);const [ptHighAmt,setPtHighAmt]=useState([]);const [ptFakeDom,setPtFakeDom]=useState([]);const [ptShowRaw,setPtShowRaw]=useState(false);
  const [ptModal,setPtModal]=useState(false);const [ptStep,setPtStep]=useState("drop");const [ptRows,setPtRows]=useState([]);const [ptLoading,setPtLoading]=useState(false);const [ptErr,setPtErr]=useState("");

  // ── Noon state ──
  const [noonRaw,setNoonRaw]=useState(null);const [adminRawFile,setAdminRawFile]=useState(null);const [noonMerged,setNoonMerged]=useState([]);const [noonEnriched,setNoonEnriched]=useState([]);const [noonFraud,setNoonFraud]=useState([]);const [noonHighAmt,setNoonHighAmt]=useState([]);const [noonFakeDom,setNoonFakeDom]=useState([]);const [noonShowRaw,setNoonShowRaw]=useState(false);const [noonLoading,setNoonLoading]=useState(false);const [noonDrag,setNoonDrag]=useState(false);const [adminDrag,setAdminDrag]=useState(false);const noonRef=useRef();const adminNoonRef=useRef();

  // ── PayMob state ──
  const [pmRaw,setPmRaw]=useState([]);const [pmSubTab,setPmSubTab]=useState("wallets");const [pmWEnriched,setPmWEnriched]=useState([]);const [pmWFraud,setPmWFraud]=useState([]);const [pmWHighAmt,setPmWHighAmt]=useState([]);const [pmWFakeDom,setPmWFakeDom]=useState([]);const [pmWalletAbusers,setPmWalletAbusers]=useState([]);const [pmBEnriched,setPmBEnriched]=useState([]);const [pmBFraud,setPmBFraud]=useState([]);const [pmBHighAmt,setPmBHighAmt]=useState([]);const [pmWShowRaw,setPmWShowRaw]=useState(false);const [pmBShowRaw,setPmBShowRaw]=useState(false);
  const [pmModal,setPmModal]=useState(false);const [pmStep,setPmStep]=useState("drop");const [pmRows,setPmRows]=useState([]);const [pmLoading,setPmLoading]=useState(false);const [pmErr,setPmErr]=useState("");

  // ── Admin state ──
  const [adminSheetRaw,setAdminSheetRaw]=useState([]);const [adminEnriched,setAdminEnriched]=useState([]);const [adminPayMethods,setAdminPayMethods]=useState([]);const [adminSuspected,setAdminSuspected]=useState([]);const [adminHighAmt,setAdminHighAmt]=useState([]);const [adminFakeDom,setAdminFakeDom]=useState([]);const [adminRechargeAbusers,setAdminRechargeAbusers]=useState([]);const [adminShowRaw,setAdminShowRaw]=useState(false);
  const [adminModal,setAdminModal]=useState(false);const [adminStep,setAdminStep]=useState("drop");const [adminRows,setAdminRows]=useState([]);const [adminLoading,setAdminLoading]=useState(false);const [adminErr,setAdminErr]=useState("");

  // ── Fawry state ──
  const [fawryRaw,setFawryRaw]=useState([]);const [fawryEnriched,setFawryEnriched]=useState([]);const [fawryHighAmt,setFawryHighAmt]=useState([]);const [fawrySuspected,setFawrySuspected]=useState([]);const [fawryFakeDom,setFawryFakeDom]=useState([]);const [fawryShowRaw,setFawryShowRaw]=useState(false);
  const [fawryModal,setFawryModal]=useState(false);const [fawryStep,setFawryStep]=useState("drop");const [fawryRows,setFawryRows]=useState([]);const [fawryLoading,setFawryLoading]=useState(false);const [fawryErr,setFawryErr]=useState("");

  // ── Promo state ──
  const [promoRaw,setPromoRaw]=useState([]);const [promoEnriched,setPromoEnriched]=useState([]);const [promoHighDiscount,setPromoHighDiscount]=useState([]);const [promoSameCard,setPromoSameCard]=useState([]);const [promoSameWallet,setPromoSameWallet]=useState([]);const [promoFakeDomain,setPromoFakeDomain]=useState([]);const [promoShowRaw,setPromoShowRaw]=useState(false);
  const [promoModal,setPromoModal]=useState(false);const [promoStep,setPromoStep]=useState("drop");const [promoRows,setPromoRows]=useState([]);const [promoLoading,setPromoLoading]=useState(false);const [promoErr,setPromoErr]=useState("");

  // ── Blocked emails state ──
  const [blockedEmails,setBlockedEmails]=useState(new Set());
  const [blockedList,setBlockedList]=useState([]);
  const [blockingEmail,setBlockingEmail]=useState(null);
  const [blockModal,setBlockModal]=useState(null); // {email, platform}
  const [blockNote,setBlockNote]=useState("");
  const [blockNoteErr,setBlockNoteErr]=useState(false);
  const [dashAlerts,setDashAlerts]=useState([]);
  const [dashSessions,setDashSessions]=useState([]);
  const [dashLoading,setDashLoading]=useState(false);
  const [dashRange,setDashRange]=useState(30);
  const [dashErr,setDashErr]=useState(null);

  // ── Import handlers ──
  const doImportPT=async()=>{await refreshBlocked();const en=ptRows.map(enrichPayTabsRow);const fraud=filterB(detectFraud(en,"CC Details")),high=filterB(detectHighAmounts(en)),fake=filterB(detectFakeDomain(en));setPtRaw(ptRows);setPtEnriched(en);setPtFraud(fraud);setPtHighAmt(high);setPtFakeDom(fake);setPtShowRaw(false);setPtStep("done");const details=`High: ${fraud.filter(r=>r.risk==="High").length} · Mid: ${fraud.filter(r=>r.risk==="Mid").length} · High Amt: ${high.length} · Fake Domain: ${fake.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"PayTabs",records:ptRows.length,details});const sessionId=await logUploadSession({platform:"paytabs",uploaded_by:currentUser.username,record_count:ptRows.length,high_count:fraud.filter(r=>r.risk==="High").length,mid_count:fraud.filter(r=>r.risk==="Mid").length,high_amt_count:high.length,fake_dom_count:fake.length});if(sessionId){await logFraudAlerts(sessionId,[...buildCCAlerts(fraud,"paytabs"),...buildHighAmtAlerts(high,"paytabs"),...buildFakeDomAlerts(fake,"paytabs")]);}await logAuditEntry({username:currentUser.username,action:"Import",platform:"PayTabs",record_count:ptRows.length,details});};
  const doImportPM=async()=>{await refreshBlocked();const wEn=pmRows.map(enrichPaymobRow);const bRows=filterBNPLRows(pmRows);const bEn=bRows.map(enrichBNPLRow);const wFraud=filterB(detectFraud(wEn,"wallets")),wHigh=filterB(detectHighAmounts(wEn)),wFake=filterB(detectFakeDomain(wEn)),wAbusers=detectWalletAbusers(wEn);const bFraud=filterB(detectBNPLFraud(bEn)),bHigh=filterB(detectHighAmounts(bEn));setPmRaw(pmRows);setPmWEnriched(wEn);setPmWFraud(wFraud);setPmWHighAmt(wHigh);setPmWFakeDom(wFake);setPmWalletAbusers(wAbusers);setPmBEnriched(bEn);setPmBFraud(bFraud);setPmBHighAmt(bHigh);setPmWShowRaw(false);setPmBShowRaw(false);setPmSubTab("wallets");setPmStep("done");const details=`W High: ${wFraud.filter(r=>r.risk==="High").length} Mid: ${wFraud.filter(r=>r.risk==="Mid").length} Abusers: ${wAbusers.length} | BNPL: Susp: ${bFraud.length} HighAmt: ${bHigh.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"PayMob",records:pmRows.length,details});const wSessionId=await logUploadSession({platform:"paymob_wallets",uploaded_by:currentUser.username,record_count:pmRows.length,high_count:wFraud.filter(r=>r.risk==="High").length,mid_count:wFraud.filter(r=>r.risk==="Mid").length,high_amt_count:wHigh.length,fake_dom_count:wFake.length,other_count:wAbusers.length});if(wSessionId){await logFraudAlerts(wSessionId,[...buildCCAlerts(wFraud,"paymob_wallets"),...buildHighAmtAlerts(wHigh,"paymob_wallets"),...buildFakeDomAlerts(wFake,"paymob_wallets"),...buildWalletAbuserAlerts(wAbusers,"paymob_wallets")]);}const bSessionId=await logUploadSession({platform:"paymob_bnpl",uploaded_by:currentUser.username,record_count:bEn.length,high_count:bFraud.length,mid_count:0,high_amt_count:bHigh.length,fake_dom_count:0});if(bSessionId){await logFraudAlerts(bSessionId,[...buildBNPLAlerts(bFraud,"paymob_bnpl"),...buildHighAmtAlerts(bHigh,"paymob_bnpl")]);}await logAuditEntry({username:currentUser.username,action:"Import",platform:"PayMob",record_count:pmRows.length,details});};
  const doImportAdmin=async()=>{await refreshBlocked();const filtered=adminRows.filter(r=>(getCol(r)("User Name")||"").toLowerCase().trim()!=="integration");const en=filtered.map(enrichAdminRow);const payM=filterB(detectAdminPayMethods(en)),susp=filterB(detectAdminSuspected(en)),high=filterB(detectAdminHighAmt(en)),fake=filterB(detectAdminFakeDomain(en)),recharge=detectRechargeAbusers(en);setAdminSheetRaw(filtered);setAdminEnriched(en);setAdminPayMethods(payM);setAdminSuspected(susp);setAdminHighAmt(high);setAdminFakeDom(fake);setAdminRechargeAbusers(recharge);setAdminShowRaw(false);setAdminStep("done");const details=`PayMethod: ${payM.length} · Suspected: ${susp.length} · High Amt: ${high.length} · Fake: ${fake.length} · Recharge: ${recharge.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"Admin",records:filtered.length,details});const sessionId=await logUploadSession({platform:"admin",uploaded_by:currentUser.username,record_count:filtered.length,high_count:payM.filter(r=>r.risk==="High").length,mid_count:payM.filter(r=>r.risk==="Mid").length,high_amt_count:high.length,fake_dom_count:fake.length,other_count:susp.length+recharge.length});if(sessionId){await logFraudAlerts(sessionId,buildAdminAlerts(payM,susp,high,fake,recharge));}await logAuditEntry({username:currentUser.username,action:"Import",platform:"Admin",record_count:filtered.length,details});};
  const doImportFawry=async()=>{await refreshBlocked();const en=fawryRows.map(enrichFawryRow);const high=filterB(detectFawryHighAmt(en)),susp=filterB(detectFawrySuspected(en)),fake=filterB(detectFawryFakeDomain(en));setFawryRaw(fawryRows);setFawryEnriched(en);setFawryHighAmt(high);setFawrySuspected(susp);setFawryFakeDom(fake);setFawryShowRaw(false);setFawryStep("done");const details=`Suspected: ${susp.length} · High Amt: ${high.length} · Fake Domain: ${fake.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"Fawry",records:fawryRows.length,details});const sessionId=await logUploadSession({platform:"fawry",uploaded_by:currentUser.username,record_count:fawryRows.length,high_count:high.length,mid_count:0,high_amt_count:high.length,fake_dom_count:fake.length,other_count:susp.length});if(sessionId){await logFraudAlerts(sessionId,buildFawryAlerts(high,susp,fake));}await logAuditEntry({username:currentUser.username,action:"Import",platform:"Fawry",record_count:fawryRows.length,details});};

  const loadNoonFile=useCallback(async(file,type)=>{try{const rows=await parseFile(file);if(type==="noon")setNoonRaw({file,rows});else setAdminRawFile({file,rows});}catch(e){alert("Cannot read: "+e.message);}},[]);
  const doNoonAnalyze=async()=>{if(!noonRaw||!adminRawFile)return;setNoonLoading(true);try{await refreshBlocked();const merged=mergeNoonAdmin(noonRaw.rows,adminRawFile.rows);const en=merged.map(enrichNoonRow);const fraud=filterB(detectFraud(en,"CC Details")),high=filterB(detectHighAmounts(en)),fake=filterB(detectFakeDomain(en));setNoonMerged(merged);setNoonEnriched(en);setNoonFraud(fraud);setNoonHighAmt(high);setNoonFakeDom(fake);setNoonShowRaw(false);const details=`High: ${fraud.filter(r=>r.risk==="High").length} · Mid: ${fraud.filter(r=>r.risk==="Mid").length} · High Amt: ${high.length} · Fake Domain: ${fake.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"Noon",records:merged.length,details});const sessionId=await logUploadSession({platform:"noon",uploaded_by:currentUser.username,record_count:merged.length,high_count:fraud.filter(r=>r.risk==="High").length,mid_count:fraud.filter(r=>r.risk==="Mid").length,high_amt_count:high.length,fake_dom_count:fake.length});if(sessionId){await logFraudAlerts(sessionId,[...buildCCAlerts(fraud,"noon"),...buildHighAmtAlerts(high,"noon"),...buildFakeDomAlerts(fake,"noon")]);}await logAuditEntry({username:currentUser.username,action:"Import",platform:"Noon",record_count:merged.length,details});}finally{setNoonLoading(false);}};
  const resetNoon=()=>{setNoonRaw(null);setAdminRawFile(null);setNoonMerged([]);setNoonEnriched([]);setNoonFraud([]);setNoonHighAmt([]);setNoonFakeDom([]);};
  const resetAdmin=()=>{setAdminSheetRaw([]);setAdminEnriched([]);setAdminPayMethods([]);setAdminSuspected([]);setAdminHighAmt([]);setAdminFakeDom([]);setAdminRechargeAbusers([]);};
  const resetFawry=()=>{setFawryRaw([]);setFawryEnriched([]);setFawryHighAmt([]);setFawrySuspected([]);setFawryFakeDom([]);};
  const doImportPromo=async()=>{await refreshBlocked();const PROMO_CARD_EXCLUDE=new Set(["VALU","TRU"]);const en=promoRows.map(enrichPromoRow).filter(r=>!PROMO_CARD_EXCLUDE.has((r._cardNum||"").toUpperCase().trim()));const highDisc=filterB(detectPromoHighDiscount(en)),sameCard=detectPromoSameCard(en),sameWallet=detectPromoSameWallet(en),fakeDom=filterB(detectPromoFakeDomain(en));setPromoRaw(promoRows);setPromoEnriched(en);setPromoHighDiscount(highDisc);setPromoSameCard(sameCard);setPromoSameWallet(sameWallet);setPromoFakeDomain(fakeDom);setPromoShowRaw(false);setPromoStep("done");const details=`High Discount: ${highDisc.length} · Same Card: ${sameCard.length} · Same Wallet: ${sameWallet.length} · Fake Domain: ${fakeDom.length}`;addAuditLog({user:currentUser.username,action:"Import",platform:"Promo",records:promoRows.length,details});const sessionId=await logUploadSession({platform:"promo",uploaded_by:currentUser.username,record_count:promoRows.length,high_count:highDisc.length,mid_count:sameCard.filter(r=>r.risk==="Mid").length+sameWallet.filter(r=>r.risk==="Mid").length,high_amt_count:0,fake_dom_count:fakeDom.length,other_count:sameCard.filter(r=>r.risk==="High").length+sameWallet.filter(r=>r.risk==="High").length});if(sessionId){await logFraudAlerts(sessionId,buildPromoAlerts(highDisc,sameCard,sameWallet,fakeDom));}await logAuditEntry({username:currentUser.username,action:"Import",platform:"Promo",record_count:promoRows.length,details});};
  const resetPromo=()=>{setPromoRaw([]);setPromoEnriched([]);setPromoHighDiscount([]);setPromoSameCard([]);setPromoSameWallet([]);setPromoFakeDomain([]);};

  const ptBadge=ptFraud.length+ptHighAmt.length+ptFakeDom.length;
  const noonBadge=noonFraud.length+noonHighAmt.length+noonFakeDom.length;
  const pmBadge=pmWFraud.length+pmWHighAmt.length+pmWFakeDom.length+pmWalletAbusers.length+pmBFraud.length+pmBHighAmt.length;
  const adminBadge=adminPayMethods.length+adminSuspected.length+adminHighAmt.length+adminFakeDom.length+adminRechargeAbusers.length;
  const fawryBadge=fawryHighAmt.length+fawrySuspected.length+fawryFakeDom.length;
  const promoBadge=promoHighDiscount.length+promoSameCard.length+promoSameWallet.length+promoFakeDomain.length;

  // ── Blocked emails — localStorage (primary) + Supabase (sync) ──
  // localStorage helpers
  const BK="wg_blocked";
  const lsGetBlocked=()=>{try{return JSON.parse(localStorage.getItem(BK)||"[]");}catch{return[];}};
  const lsSaveBlocked=(list)=>{try{localStorage.setItem(BK,JSON.stringify(list));}catch{}};

  const refreshBlocked=async()=>{
    // Always start from localStorage (guaranteed to work)
    const local=lsGetBlocked();
    setBlockedList(local);
    setBlockedEmails(new Set(local.map(r=>r.entity_value)));
    // Try to merge remote entries (best-effort, don't wipe local on failure)
    try{
      const remote=await loadBlockedEmails();
      if(remote&&remote.length>0){
        const merged=[...local];
        remote.forEach(r=>{if(!merged.find(m=>m.entity_value===r.entity_value))merged.push(r);});
        lsSaveBlocked(merged);
        setBlockedList(merged);
        setBlockedEmails(new Set(merged.map(r=>r.entity_value)));
      }
    }catch(e){/* Supabase unavailable — localStorage is source of truth */}
  };
  useEffect(()=>{if(currentUser)refreshBlocked();},[currentUser?.username]);

  const loadDash=async(days)=>{
    setDashLoading(true);setDashErr(null);
    try{
      const {alerts,sessions}=await loadDashboardStats(days);
      setDashAlerts(alerts);setDashSessions(sessions);
    }catch(e){setDashErr("Could not load dashboard data. Check Supabase RLS settings.");}
    finally{setDashLoading(false);}
  };
  useEffect(()=>{if(tab==="dashboard"&&currentUser)loadDash(dashRange);},[tab,dashRange,currentUser?.username]);

  if(!currentUser)return <LoginScreen onLogin={setCurrentUser}/>;

  const isBlocked=(email)=>!!email&&blockedEmails.has(email.toLowerCase().trim());

  const handleBlock=async(email,platform,note)=>{
    if(!email)return;
    const key=email.toLowerCase().trim();
    const entry={entity_value:key,blocked_by:currentUser.username,platform:platform||null,note:note||null,created_at:new Date().toISOString()};
    // 1. Write localStorage immediately — this is the source of truth
    const current=lsGetBlocked();
    const updated=[entry,...current.filter(b=>b.entity_value!==key)];
    lsSaveBlocked(updated);
    // 2. Update React state immediately — card disappears now
    setBlockedEmails(prev=>new Set([...prev,key]));
    setBlockedList(updated);
    setBlockingEmail(key);
    // 3. Try Supabase in background — best effort, never wipes local state
    try{await blockEmail(key,currentUser.username,platform||"unknown",note||undefined);}catch(e){console.error("blockEmail:",e);}
    setBlockingEmail(null);
  };

  // Opens the block reason modal instead of blocking immediately
  const openBlockModal=(email,platform)=>{setBlockModal({email,platform});setBlockNote("");setBlockNoteErr(false);};
  const confirmBlock=async()=>{
    if(!blockModal)return;
    if(!blockNote.trim()){setBlockNoteErr(true);return;}
    await handleBlock(blockModal.email,blockModal.platform,blockNote.trim());
    setBlockModal(null);setBlockNote("");setBlockNoteErr(false);
  };
  const handleUnblock=async(email)=>{
    const key=email.toLowerCase().trim();
    // 1. Remove from localStorage immediately
    const updated=lsGetBlocked().filter(b=>b.entity_value!==key);
    lsSaveBlocked(updated);
    // 2. Update React state immediately
    setBlockedEmails(prev=>{const n=new Set(prev);n.delete(key);return n;});
    setBlockedList(updated);
    // 3. Try Supabase in background
    try{await unblockEmail(key);}catch(e){console.error("unblockEmail:",e);}
  };

  // Filters an array of email-centric results against the current blocked set
  const filterB=(arr)=>(arr||[]).filter(r=>!blockedEmails.has((r.email||"").toLowerCase().trim()));

  // ── Shared block props passed into every dashboard config ──
  const blockProps={onBlock:openBlockModal,isBlocked,blockingEmail};
  const unblocked=(arr)=>(arr||[]).filter(r=>!isBlocked(r.email));
  const unblockedGroup=(arr)=>(arr||[]).map(r=>({...r,emails:(r.emails||[]).filter(e=>!isBlocked(e))})).filter(r=>(r.emails||[]).length>=2);

  // ── PayTabs dashboard config ──
  const ptConfig={
    accent:"#7c3aed",reimportLabel:"PayTabs",raw:ptRaw,enriched:ptEnriched,
    stats:[{label:"Total Records",value:ptRaw.length,clr:"#7c3aed",Icon:CreditCard},{label:"High Risk",value:ptFraud.filter(r=>r.risk==="High").length,clr:"#dc2626",Icon:AlertTriangle},{label:"Mid Risk",value:ptFraud.filter(r=>r.risk==="Mid").length,clr:"#d97706",Icon:Users},{label:"High Amounts",value:ptHighAmt.length,clr:"#065f46",Icon:DollarSign},{label:"Fake Domain",value:ptFakeDom.length,clr:"#7c3aed",Icon:ShieldAlert}],
    tabs:[
      {id:"all",label:"All",rows:unblocked(ptFraud),accent:"#334155",cfg:{ccLabel:"CC Details",cartLabel:"Cart ID",showOS:false,showECI:true,showCountries:true,hideStatus:true}},
      {id:"high",label:"🔴 High",rows:unblocked(ptFraud.filter(r=>r.risk==="High")),accent:"#dc2626",cfg:{ccLabel:"CC Details",cartLabel:"Cart ID",showOS:false,showECI:true,showCountries:true,hideStatus:true}},
      {id:"mid",label:"🟡 Mid",rows:unblocked(ptFraud.filter(r=>r.risk==="Mid")),accent:"#d97706",cfg:{ccLabel:"CC Details",cartLabel:"Cart ID",showOS:false,showECI:true,showCountries:true,hideStatus:true}},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(ptHighAmt),accent:"#065f46",cfg:{ccLabel:"CC Details",cartLabel:"Cart ID",showOS:false,showECI:true,showCountries:true,hideStatus:true}},
      {id:"fakedomain",label:"📧 Fake Domain",rows:unblocked(ptFakeDom),accent:"#7c3aed",cfg:{ccLabel:"CC Details",cartLabel:"Cart ID",showOS:false,showECI:true,showCountries:true,hideStatus:true}},
    ],
    onReimport:()=>{setPtModal(true);setPtStep("drop");setPtRows([]);setPtErr("");},
    showRaw:ptShowRaw,setShowRaw:setPtShowRaw,
    previewCols:["Customer Email","Customer Name","Payment Description","Expiry Year","Expiry Month","Issuer Country","Cart ID"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>makeExcelFile(ptFraud,ptHighAmt,ptFakeDom,"PayTabs"),
    ...blockProps,
  };

  // ── Noon dashboard config ──
  const noonConfig={
    accent:"#f59e0b",reimportLabel:"Noon",raw:noonMerged,enriched:noonEnriched,
    stats:[{label:"Total Records",value:noonMerged.length,clr:"#f59e0b",Icon:CreditCard},{label:"High Risk",value:noonFraud.filter(r=>r.risk==="High").length,clr:"#dc2626",Icon:AlertTriangle},{label:"Mid Risk",value:noonFraud.filter(r=>r.risk==="Mid").length,clr:"#d97706",Icon:Users},{label:"High Amounts",value:noonHighAmt.length,clr:"#065f46",Icon:DollarSign},{label:"Fake Domain",value:noonFakeDom.length,clr:"#7c3aed",Icon:ShieldAlert}],
    tabs:[
      {id:"all",label:"All",rows:unblocked(noonFraud),accent:"#334155",cfg:{ccLabel:"CC Details",cartLabel:"Merchant Order Ref",showOS:true,showECI:true,showCountries:true}},
      {id:"high",label:"🔴 High",rows:unblocked(noonFraud.filter(r=>r.risk==="High")),accent:"#dc2626",cfg:{ccLabel:"CC Details",cartLabel:"Merchant Order Ref",showOS:true,showECI:true,showCountries:true}},
      {id:"mid",label:"🟡 Mid",rows:unblocked(noonFraud.filter(r=>r.risk==="Mid")),accent:"#d97706",cfg:{ccLabel:"CC Details",cartLabel:"Merchant Order Ref",showOS:true,showECI:true,showCountries:true}},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(noonHighAmt),accent:"#065f46",cfg:{ccLabel:"CC Details",cartLabel:"Merchant Order Ref",showOS:true,showECI:true,showCountries:true}},
      {id:"fakedomain",label:"📧 Fake Domain",rows:unblocked(noonFakeDom),accent:"#7c3aed",cfg:{ccLabel:"CC Details",cartLabel:"Merchant Order Ref",showOS:true,showECI:true,showCountries:true}},
    ],
    onReimport:resetNoon,
    showRaw:noonShowRaw,setShowRaw:setNoonShowRaw,
    previewCols:["User Email","User Name","Payerinfo","Merchantorderreference","Orderdate_UTC","Issuercountry","Responsemessage","Orderstatus"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>makeExcelFile(noonFraud,noonHighAmt,noonFakeDom,"Noon"),
    ...blockProps,
  };

  // ── PayMob wallets dashboard config ──
  const pmWConfig={
    accent:"#2563eb",reimportLabel:"PayMob Wallets",raw:pmRaw,enriched:pmWEnriched,
    stats:[{label:"Wallet Records",value:pmRaw.length,clr:"#2563eb",Icon:CreditCard},{label:"High Risk",value:pmWFraud.filter(r=>r.risk==="High").length,clr:"#dc2626",Icon:AlertTriangle},{label:"Mid Risk",value:pmWFraud.filter(r=>r.risk==="Mid").length,clr:"#d97706",Icon:Users},{label:"High Amounts",value:pmWHighAmt.length,clr:"#065f46",Icon:DollarSign},{label:"Fake Domain",value:pmWFakeDom.length,clr:"#7c3aed",Icon:ShieldAlert},{label:"Wallet Abusers",value:pmWalletAbusers.length,clr:"#7c3aed",Icon:ShieldAlert}],
    tabs:[
      {id:"all",label:"All",rows:unblocked(pmWFraud),accent:"#334155",cfg:{ccLabel:"Wallets",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"high",label:"🔴 High",rows:unblocked(pmWFraud.filter(r=>r.risk==="High")),accent:"#dc2626",cfg:{ccLabel:"Wallets",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"mid",label:"🟡 Mid",rows:unblocked(pmWFraud.filter(r=>r.risk==="Mid")),accent:"#d97706",cfg:{ccLabel:"Wallets",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(pmWHighAmt),accent:"#065f46",cfg:{ccLabel:"Wallets",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"fakedomain",label:"📧 Fake Domain",rows:unblocked(pmWFakeDom),accent:"#7c3aed",cfg:{ccLabel:"Wallets",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"walletabuser",label:"👛 Wallet Abusers",rows:unblockedGroup(pmWalletAbusers),accent:"#7c3aed",renderCard:(r,i,ac)=><WalletAbuserCard key={i} r={r} accent={ac}/>},
    ],
    onReimport:()=>{setPmModal(true);setPmStep("drop");setPmRows([]);setPmErr("");},
    showRaw:pmWShowRaw,setShowRaw:setPmWShowRaw,
    previewCols:["client_email","client_name","client_phone","merchant_order_id","created_at","data_message_execl","success","amount_whole"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>makeExcelFile(pmWFraud,pmWHighAmt,pmWFakeDom,"PayMob_Wallets"),
    ...blockProps,
  };

  // ── PayMob BNPL config ──
  const pmBConfig={
    accent:"#7c3aed",reimportLabel:"PayMob BNPL",raw:pmBEnriched,enriched:pmBEnriched,
    stats:[{label:"BNPL Records",value:pmBEnriched.length,clr:"#7c3aed",Icon:CreditCard},{label:"High Suspicious",value:pmBFraud.length,clr:"#dc2626",Icon:AlertTriangle},{label:"High Amounts",value:pmBHighAmt.length,clr:"#065f46",Icon:DollarSign}],
    tabs:[
      {id:"all",label:"🔴 Suspicious",rows:unblocked(pmBFraud),accent:"#dc2626",cfg:{ccLabel:"BNPL",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(pmBHighAmt),accent:"#065f46",cfg:{ccLabel:"BNPL",cartLabel:"Merchant Order ID",showOS:true,showECI:false,showCountries:false}},
    ],
    onReimport:()=>{setPmModal(true);setPmStep("drop");setPmRows([]);setPmErr("");},
    showRaw:pmBShowRaw,setShowRaw:setPmBShowRaw,
    previewCols:["client_email","client_name","payment_method","merchant_order_id","created_at","success","amount_whole"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>makeExcelFile(pmBFraud,pmBHighAmt,[],"PayMob_BNPL"),
    ...blockProps,
  };

  // ── Admin dashboard config ──
  // Admin: build ccTxRows using real sheet fields (no authMsg, use timestamp/status/amt from sheet)
  const makeAdminTxRows=rows=>(rows||[]).map(rx=>({cartId:rx._txId||"",timestamp:rx._timestamp||"",authMsg:null,orderStatus:rx._status||"",amt:rx._amt||0,biller:rx._biller||"",payMethod:rx._payMethod||""}));
  const adminSuspectedWithTxRows=adminSuspected.map(r=>({...r,uniqueCCs:[r.userId||"—"],ccTxRows:{[r.userId||"—"]:makeAdminTxRows(r.rows)},email:r.email||"",custNames:r.custNames||[]}));
  const adminPayMethodsAdapted=adminPayMethods.map(r=>({...r,uniqueCCs:r.uniqueMethods||[],ccTxRows:Object.fromEntries((r.uniqueMethods||[]).map(m=>[m,makeAdminTxRows((r.rows||[]).filter(rx=>rx._payMethod===m))]))}));
  const adminHighAmtAdapted=adminHighAmt.map(r=>({...r,uniqueCCs:r.uniqueMethods&&r.uniqueMethods.length>0?r.uniqueMethods:["—"],ccTxRows:r.uniqueMethods&&r.uniqueMethods.length>0?Object.fromEntries(r.uniqueMethods.map(m=>[m,makeAdminTxRows(r.rows.filter(rx=>rx._payMethod===m))])):{"—":makeAdminTxRows(r.rows)}}));
  const adminFakeDomAdapted=adminFakeDom.map(r=>({...r,uniqueCCs:r.uniqueMethods&&r.uniqueMethods.length>0?r.uniqueMethods:["—"],ccTxRows:r.uniqueMethods&&r.uniqueMethods.length>0?Object.fromEntries(r.uniqueMethods.map(m=>[m,makeAdminTxRows(r.rows.filter(rx=>rx._payMethod===m))])):{"—":makeAdminTxRows(r.rows)}}));
  const adminRechargeAdapted=adminRechargeAbusers.map(r=>({...r,uniqueCCs:[r.recharge||"—"],ccTxRows:{[r.recharge||"—"]:makeAdminTxRows(r.rows)},email:r.emails.join(", ")||"",custNames:[]}));

  const adminConfig={
    accent:"#0ea5e9",reimportLabel:"Admin",raw:adminSheetRaw,enriched:adminEnriched,
    stats:[{label:"Total Records",value:adminSheetRaw.length,clr:"#0ea5e9",Icon:CreditCard},{label:"Pay Method Abuse",value:adminPayMethods.length,clr:"#dc2626",Icon:AlertTriangle},{label:"Suspected",value:adminSuspected.length,clr:"#7c3aed",Icon:Users},{label:"High Amounts",value:adminHighAmt.length,clr:"#065f46",Icon:DollarSign},{label:"Fake Domain",value:adminFakeDom.length,clr:"#581c87",Icon:ShieldAlert},{label:"Recharge Abusers",value:adminRechargeAbusers.length,clr:"#7c3aed",Icon:ShieldAlert}],
    tabs:[
      {id:"paymethods",label:"💳 Pay Method Abuse",rows:unblocked(adminPayMethodsAdapted),accent:"#dc2626",cfg:{ccLabel:"Payment Method",cartLabel:"Transaction ID",showOS:true,showECI:false,showCountries:false,hideAuthMsg:true,showBiller:true}},
      {id:"suspected",label:"🕵️ Suspected",rows:unblocked(adminSuspectedWithTxRows),accent:"#7c3aed",cfg:{ccLabel:"User ID",cartLabel:"Transaction ID",showOS:true,showECI:false,showCountries:false,hideAuthMsg:true,showBiller:true,showPayMethod:true}},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(adminHighAmtAdapted),accent:"#065f46",cfg:{ccLabel:"Payment Method",cartLabel:"Transaction ID",showOS:true,showECI:false,showCountries:false,hideAuthMsg:true,showBiller:true}},
      {id:"fakedomain",label:"📧 Fake Domain",rows:unblocked(adminFakeDomAdapted),accent:"#7c3aed",cfg:{ccLabel:"Payment Method",cartLabel:"Transaction ID",showOS:true,showECI:false,showCountries:false,hideAuthMsg:true,showBiller:true}},
      {id:"rechargeabuser",label:"📱 Recharge Abusers",rows:adminRechargeAdapted,accent:"#7c3aed",cfg:{ccLabel:"Recharge Number",cartLabel:"Transaction ID",showOS:true,showECI:false,showCountries:false,hideAuthMsg:true,showBiller:true,showPayMethod:true}},
    ],
    onReimport:()=>{resetAdmin();setAdminModal(true);setAdminStep("drop");setAdminRows([]);setAdminErr("");},
    showRaw:adminShowRaw,setShowRaw:setAdminShowRaw,
    previewCols:["User Email","User Name","Payment Method","Service Biller","Status","Requested Amount","Transaction create time"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>{const wb=XLSX.utils.book_new();[["PayMethods",adminPayMethodsAdapted],["Suspected",adminSuspectedWithTxRows],["HighAmounts",adminHighAmtAdapted],["FakeDomain",adminFakeDomAdapted],["Recharge",adminRechargeAdapted]].forEach(([name,rows])=>{const data=rows.map(r=>({"Email":r.email||"—","Names":(r.custNames||[]).join(", ")||"—","CC/Method/ID":(r.uniqueCCs||[]).join(", ")||"—","Total":r.totalAmt||0,"Reason":(r.reasons||[]).join(" | ")}));const ws=XLSX.utils.json_to_sheet(data.length>0?data:[{}]);XLSX.utils.book_append_sheet(wb,ws,name);});XLSX.writeFile(wb,`Admin_Fraud_${new Date().toISOString().split("T")[0]}.xlsx`);},
    ...blockProps,
  };

  // ── Fawry dashboard config ──
  const fawryAllTxRows=r=>({...r,uniqueCCs:[r.email||"—"],ccTxRows:{[r.email||"—"]:r.txRows.map(tx=>({cartId:tx.txId||"",fawryCode:tx.fawryCode||"",timestamp:undefined,soldDate:tx.soldDate||"",authMsg:undefined,orderStatus:tx.status||"",amt:tx.amt||0}))}});
  const fawrySuspAdapted=fawrySuspected.map(fawryAllTxRows);
  const fawryHighAdapted=fawryHighAmt.map(fawryAllTxRows);
  const fawryFakeAdapted=fawryFakeDom.map(fawryAllTxRows);

  // Fawry uses custom table rows since fields differ
  const FawryCard=({r,accentColor})=>(
    <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:`1px solid ${accentColor}22`,marginBottom:16,overflow:"hidden"}}>
      <div style={{display:"flex",alignItems:"stretch"}}>
        <div style={{background:accentColor,width:80,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:2}}>
          <div style={{color:"#fff",fontWeight:900,fontSize:11,letterSpacing:1,textAlign:"center"}}>SUSPECTED</div>
          <div style={{color:"rgba(255,255,255,0.6)",fontSize:10,marginTop:6,fontWeight:600}}>{r.txCount} tx</div>
        </div>
        <div style={{flex:1,padding:"16px 20px",borderRight:"1px solid #f1f5f9"}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:5}}>{(r.custNames||[]).map((n,i)=><span key={i} style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>{n}</span>)}</div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:13,color:"#334155",wordBreak:"break-all"}}>{r.email}</span>
            {isDisposable(r.email)&&<span style={{background:"#fee2e2",color:"#b91c1c",padding:"2px 7px",borderRadius:4,fontSize:11,fontWeight:700}}>⚠ Non-wl</span>}
          </div>
        </div>
        <div style={{background:"#f8fafc",padding:"16px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:130,borderRight:"1px solid #f1f5f9"}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Total (EGP)</div>
          <div style={{fontWeight:900,color:accentColor,fontSize:24,lineHeight:1}}>{(r.totalAmt||0).toLocaleString()}</div>
        </div>
        <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220,maxWidth:280}}>
          <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>Reason</div>
          {(r.reasons||[]).map((rs,i)=><div key={i} style={{display:"flex",gap:6,marginBottom:4}}><span style={{color:accentColor,fontWeight:900,fontSize:14}}>›</span><span style={{fontSize:12,color:"#475569",lineHeight:1.5}}>{rs}</span></div>)}
        </div>
      </div>
      {(r.txRows||[]).length>0&&(
        <div style={{borderTop:"1px solid #f1f5f9",overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:"#f8fafc"}}>{["Transaction ID","Fawry Code","Sold Date","Status","Amount (EGP)"].map(h=><th key={h} style={{padding:"8px 14px",textAlign:"left",fontWeight:700,color:"#64748b",fontSize:10,textTransform:"uppercase",letterSpacing:0.6,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead>
            <tbody>{(r.txRows||[]).map((tx,ti)=>{const sl=(tx.status||"").toLowerCase();const clr=sl==="paid"?"#16a34a":sl==="refunded"?"#d97706":"#64748b";return(<tr key={ti} style={{borderBottom:ti<(r.txRows||[]).length-1?"1px solid #f8fafc":"none",background:ti%2===0?"#fff":"#fafafa"}}>
              <td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{tx.txId?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{tx.txId}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td>
              <td style={{padding:"9px 14px",fontFamily:"monospace",fontWeight:700,color:"#a5b4fc",fontSize:12,whiteSpace:"nowrap"}}>{tx.fawryCode||"—"}</td>
              <td style={{padding:"9px 14px",fontSize:11,fontFamily:"monospace",color:"#334155",background:"#f8fafc",whiteSpace:"nowrap"}}>{tx.soldDate||"—"}</td>
              <td style={{padding:"9px 14px",whiteSpace:"nowrap"}}><span style={{fontWeight:700,color:clr,fontSize:12}}>{tx.status||"—"}</span></td>
              <td style={{padding:"9px 14px",textAlign:"right",fontWeight:700}}>{tx.amt>0?<span style={{background:"#f0fdf4",color:"#16a34a",padding:"2px 8px",borderRadius:5,fontSize:12}}>{tx.amt.toLocaleString()}</span>:"—"}</td>
            </tr>);})}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  // ── Promo Cards ──
  const PromoCard=({r,accent,badgeLabel,onBlock,isBlocked,blockingEmail})=>{const badgeBg=badgeLabel==="FAKE DOMAIN"?"#581c87":"#dc2626";const bLines=badgeLabel.split(" ");return(<div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:`1px solid ${accent}22`,marginBottom:16,overflow:"hidden"}}><div style={{display:"flex",alignItems:"stretch"}}><div style={{background:badgeBg,width:80,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:2}}>{bLines.map((t,i)=><div key={i} style={{color:"#fff",fontWeight:900,fontSize:10,letterSpacing:1,textAlign:"center"}}>{t}</div>)}<div style={{color:"rgba(255,255,255,0.6)",fontSize:10,marginTop:6,fontWeight:600}}>{r.txCount} orders</div></div><div style={{flex:1,padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",borderRight:"1px solid #f1f5f9"}}><div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:5}}>{(r.custNames||[]).map((n,i)=><span key={i} style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>{n}</span>)}</div><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}><span style={{fontSize:13,color:"#334155",wordBreak:"break-all"}}>{r.email}</span>{r.email&&isDisposable(r.email)&&<span style={{background:"#fee2e2",color:"#b91c1c",padding:"2px 7px",borderRadius:4,fontSize:11,fontWeight:700}}>⚠ Non-wl</span>}</div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{(r.promoCodes||[]).map((code,i)=><div key={i} style={{background:"#1e293b",color:"#fbbf24",borderRadius:6,padding:"5px 11px",fontSize:12,fontFamily:"'Courier New',monospace",fontWeight:700,display:"inline-flex",alignItems:"center",gap:6}}><span style={{fontSize:11}}>🏷</span>{code}</div>)}</div></div><div style={{background:"#f8fafc",padding:"16px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:130,borderRight:"1px solid #f1f5f9"}}><div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Discount (EGP)</div><div style={{fontWeight:900,color:accent,fontSize:24,lineHeight:1}}>{(r.totalDiscount||0).toLocaleString()}</div></div><div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220,maxWidth:280}}><div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>Reason</div>{(r.reasons||[]).map((rs,i)=><div key={i} style={{display:"flex",gap:6,marginBottom:4}}><span style={{color:accent,fontWeight:900,fontSize:14}}>›</span><span style={{fontSize:12,color:"#475569",lineHeight:1.5}}>{rs}</span></div>)}</div>{onBlock&&r.email&&<div style={{padding:"16px 14px",display:"flex",alignItems:"center",justifyContent:"center",borderLeft:"1px solid #f1f5f9"}}><button onClick={()=>onBlock(r.email,"promo")} disabled={!!blockingEmail} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4,padding:"10px 14px",background:isBlocked&&isBlocked(r.email)?"#fef2f2":"#fff7ed",border:`1.5px solid ${isBlocked&&isBlocked(r.email)?"#fca5a5":"#fed7aa"}`,borderRadius:10,cursor:"pointer",opacity:blockingEmail===r.email?.toLowerCase()?.trim()?0.5:1}}><span style={{fontSize:16}}>{isBlocked&&isBlocked(r.email)?"🔒":"🚫"}</span><span style={{fontSize:10,fontWeight:700,color:isBlocked&&isBlocked(r.email)?"#dc2626":"#ea580c",whiteSpace:"nowrap"}}>{isBlocked&&isBlocked(r.email)?"Blocked":"Block"}</span></button></div>}</div>{(r.txRows||[]).length>0&&(<div style={{borderTop:"1px solid #f1f5f9",overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr style={{background:"#f8fafc"}}>{["Order ID","Promo Code","Trx Flag","Discount (EGP)","Status","Total (EGP)","Date"].map(h=><th key={h} style={{padding:"8px 14px",textAlign:"left",fontWeight:700,color:"#64748b",fontSize:10,textTransform:"uppercase",letterSpacing:0.6,whiteSpace:"nowrap",borderBottom:"1px solid #e2e8f0"}}>{h}</th>)}</tr></thead><tbody>{(r.txRows||[]).map((tx,ti)=>{const sl=(tx.orderStatus||"").toLowerCase();const clr=sl==="paid"||sl==="completed"?"#16a34a":sl==="refunded"?"#d97706":"#64748b";return(<tr key={ti} style={{borderBottom:ti<(r.txRows||[]).length-1?"1px solid #f8fafc":"none",background:ti%2===0?"#fff":"#fafafa"}}><td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{tx.orderId?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{tx.orderId}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td><td style={{padding:"9px 14px",whiteSpace:"nowrap"}}><span style={{background:"#1e293b",color:"#fbbf24",padding:"2px 9px",borderRadius:5,fontSize:11,fontFamily:"monospace",fontWeight:700}}>{tx.promoCode||"—"}</span></td><td style={{padding:"9px 14px",whiteSpace:"nowrap"}}>{tx.trxFlag?<span style={{fontWeight:700,fontSize:12,color:tx.trxFlag.toLowerCase()==="fraud"?"#dc2626":tx.trxFlag.toLowerCase()==="abuser"?"#d97706":"#64748b"}}>{tx.trxFlag}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td><td style={{padding:"9px 14px",textAlign:"right",fontWeight:700,color:accent}}>{tx.discountAmt>0?tx.discountAmt.toLocaleString():"—"}</td><td style={{padding:"9px 14px",whiteSpace:"nowrap"}}><span style={{fontWeight:700,color:clr,fontSize:12}}>{tx.orderStatus||"—"}</span></td><td style={{padding:"9px 14px",textAlign:"right",fontWeight:700}}>{tx.total>0?<span style={{background:"#f0fdf4",color:"#16a34a",padding:"2px 8px",borderRadius:5,fontSize:12}}>{tx.total.toLocaleString()}</span>:"—"}</td><td style={{padding:"9px 14px",fontSize:11,fontFamily:"monospace",color:"#475569",whiteSpace:"nowrap"}}>{tx.createdAt||"—"}</td></tr>);})}</tbody></table></div>)}</div>);};
  const PromoAbuserCard=({r,accent,identLabel})=>{const isH=r.risk==="High";const identifier=r.card||r.wallet;return(<div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:`1px solid ${accent}22`,marginBottom:16,overflow:"hidden"}}><div style={{display:"flex",alignItems:"stretch"}}><div style={{background:isH?"#dc2626":"#d97706",width:80,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:2}}><div style={{color:"#fff",fontWeight:900,fontSize:11,letterSpacing:1}}>{r.risk.toUpperCase()}</div><div style={{color:"rgba(255,255,255,0.6)",fontSize:10,marginTop:6,fontWeight:600}}>{r.txCount} tx</div></div><div style={{flex:1,padding:"16px 20px",borderRight:"1px solid #f1f5f9"}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}><div style={{background:"#1e293b",color:"#a5b4fc",borderRadius:7,padding:"6px 14px",fontSize:13,fontFamily:"monospace",fontWeight:700}}>{identifier}</div><span style={{background:"#fef3c7",color:"#92400e",padding:"3px 10px",borderRadius:6,fontSize:11,fontWeight:700}}>{identLabel} · 👥 {r.emails.length} accounts</span></div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{r.emails.map((email,ei)=><span key={ei} style={{fontSize:12,color:"#334155",background:"#f8fafc",borderRadius:5,padding:"3px 9px",border:"1px solid #e2e8f0"}}>{email}{isDisposable(email)&&<span style={{marginLeft:5,background:"#fee2e2",color:"#b91c1c",padding:"1px 5px",borderRadius:3,fontSize:10,fontWeight:600}}>⚠</span>}</span>)}</div></div><div style={{background:"#f8fafc",padding:"16px 24px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:130,borderRight:"1px solid #f1f5f9"}}><div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Discount (EGP)</div><div style={{fontWeight:900,color:accent,fontSize:24,lineHeight:1}}>{r.totalDiscount.toLocaleString()}</div></div><div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220}}><div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:8}}>Reason</div>{r.reasons.map((rs,ri)=><div key={ri} style={{display:"flex",gap:6,marginBottom:4}}><span style={{color:accent,fontWeight:900,fontSize:14}}>›</span><span style={{fontSize:12,color:"#475569",lineHeight:1.5}}>{rs}</span></div>)}</div></div>{r.emails.map((email,ei)=>{const d=r.emailDetails[email]||{};return(<div key={ei} style={{borderTop:"1px solid #f1f5f9"}}><div style={{padding:"8px 16px",background:"#f8fafc",display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:11,fontWeight:700,color:"#475569"}}>📧</span><span style={{fontSize:12,color:"#334155",fontWeight:600}}>{email}</span><span style={{fontSize:11,color:"#94a3b8"}}>· {(d.txRows||[]).length} orders · Discount EGP {(d.discount||0).toLocaleString()}</span></div>{(d.txRows||[]).length>0&&(<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead><tr style={{background:"#fff"}}>{["Order ID","Promo Code","Trx Flag","Discount (EGP)","Status","Date"].map(h=><th key={h} style={{padding:"6px 14px",textAlign:"left",fontWeight:700,color:"#94a3b8",fontSize:10,textTransform:"uppercase",letterSpacing:0.5,borderBottom:"1px solid #f1f5f9",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead><tbody>{(d.txRows||[]).map((tx,ti)=>{const sl=(tx.orderStatus||"").toLowerCase();const clr=sl==="paid"||sl==="completed"?"#16a34a":sl==="refunded"?"#d97706":"#64748b";return(<tr key={ti} style={{borderBottom:ti<(d.txRows||[]).length-1?"1px solid #f8fafc":"none",background:ti%2===0?"#fff":"#fafafa"}}><td style={{padding:"8px 14px",whiteSpace:"nowrap"}}>{tx.orderId?<span style={{background:"#eff6ff",color:"#2563eb",border:"1px solid #bfdbfe",padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{tx.orderId}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td><td style={{padding:"8px 14px",whiteSpace:"nowrap"}}><span style={{background:"#1e293b",color:"#fbbf24",padding:"2px 9px",borderRadius:5,fontSize:11,fontFamily:"monospace",fontWeight:700}}>{tx.promoCode||"—"}</span></td><td style={{padding:"8px 14px",whiteSpace:"nowrap"}}>{tx.trxFlag?<span style={{fontWeight:700,fontSize:12,color:tx.trxFlag.toLowerCase()==="fraud"?"#dc2626":tx.trxFlag.toLowerCase()==="abuser"?"#d97706":"#64748b"}}>{tx.trxFlag}</span>:<span style={{color:"#cbd5e1"}}>—</span>}</td><td style={{padding:"8px 14px",textAlign:"right",fontWeight:700,color:accent}}>{tx.discountAmt>0?tx.discountAmt.toLocaleString():"—"}</td><td style={{padding:"8px 14px",whiteSpace:"nowrap"}}><span style={{fontWeight:700,color:clr,fontSize:12}}>{tx.orderStatus||"—"}</span></td><td style={{padding:"8px 14px",fontSize:11,fontFamily:"monospace",color:"#475569",whiteSpace:"nowrap"}}>{tx.createdAt||"—"}</td></tr>);})}</tbody></table>)}</div>);})}</div>);};

  const fawryConfig={
    accent:"#f97316",reimportLabel:"Fawry",raw:fawryRaw,enriched:fawryEnriched,
    stats:[{label:"Total Records",value:fawryRaw.length,clr:"#f97316",Icon:CreditCard},{label:"Suspected",value:fawrySuspected.length,clr:"#dc2626",Icon:AlertTriangle},{label:"High Amounts",value:fawryHighAmt.length,clr:"#065f46",Icon:DollarSign},{label:"Fake Domain",value:fawryFakeDom.length,clr:"#7c3aed",Icon:ShieldAlert}],
    tabs:[
      {id:"suspected",label:"🔴 Suspected",rows:unblocked(fawrySuspected),accent:"#dc2626",renderCard:(r,i,ac)=><FawryCard key={i} r={r} accentColor={ac}/>},
      {id:"highamount",label:"💰 High Amounts",rows:unblocked(fawryHighAmt),accent:"#065f46",renderCard:(r,i,ac)=><FawryCard key={i} r={{...r,reasons:r.reasons}} accentColor={ac}/>},
      {id:"fakedomain",label:"📧 Fake Domain",rows:unblocked(fawryFakeDom),accent:"#7c3aed",renderCard:(r,i,ac)=><FawryCard key={i} r={r} accentColor={ac}/>},
    ],
    onReimport:()=>{resetFawry();setFawryModal(true);setFawryStep("drop");setFawryRows([]);setFawryErr("");},
    showRaw:fawryShowRaw,setShowRaw:setFawryShowRaw,
    previewCols:["user","user email","Requested","Fawry Code","Status","Payment Method","Sold Date"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>{const wb=XLSX.utils.book_new();[["Suspected",fawrySuspected],["HighAmounts",fawryHighAmt],["FakeDomain",fawryFakeDom]].forEach(([name,rows])=>{const data=rows.map(r=>({"Email":r.email||"—","Names":(r.custNames||[]).join(", ")||"—","Tx Count":r.txCount,"Total (EGP)":r.totalAmt||0,"Reason":(r.reasons||[]).join(" | ")}));const ws=XLSX.utils.json_to_sheet(data.length>0?data:[{}]);XLSX.utils.book_append_sheet(wb,ws,name);});XLSX.writeFile(wb,`Fawry_Fraud_${new Date().toISOString().split("T")[0]}.xlsx`);},
    ...blockProps,
  };

  // ── Promo dashboard config ──
  const promoConfig={
    accent:"#e11d48",reimportLabel:"Promo",raw:promoRaw,enriched:promoEnriched,
    stats:[{label:"Total Records",value:promoRaw.length,clr:"#e11d48",Icon:CreditCard},{label:"High Discount",value:promoHighDiscount.length,clr:"#dc2626",Icon:AlertTriangle},{label:"Same Card",value:promoSameCard.length,clr:"#d97706",Icon:Users},{label:"Same Wallet",value:promoSameWallet.length,clr:"#7c3aed",Icon:ShieldAlert},{label:"Fake Domain",value:promoFakeDomain.length,clr:"#581c87",Icon:ShieldAlert}],
    tabs:[
      {id:"highdiscount",label:"💸 High Discount",rows:promoHighDiscount.filter(r=>!isBlocked(r.email)),accent:"#dc2626",renderCard:(r,i,ac)=><PromoCard key={i} r={r} accent={ac} badgeLabel="HIGH DISCOUNT" onBlock={e=>openBlockModal(e,"promo")} isBlocked={isBlocked} blockingEmail={blockingEmail}/>},
      {id:"samecard",label:"💳 Same Card",rows:promoSameCard.map(r=>({...r,emails:r.emails.filter(e=>!isBlocked(e))})).filter(r=>r.emails.length>=2),accent:"#d97706",renderCard:(r,i,ac)=><PromoAbuserCard key={i} r={r} accent={ac} identLabel="Card"/>},
      {id:"samewallet",label:"👛 Same Wallet",rows:promoSameWallet.map(r=>({...r,emails:r.emails.filter(e=>!isBlocked(e))})).filter(r=>r.emails.length>=2),accent:"#7c3aed",renderCard:(r,i,ac)=><PromoAbuserCard key={i} r={r} accent={ac} identLabel="Wallet"/>},
      {id:"fakedomain",label:"📧 Fake Domain",rows:promoFakeDomain.filter(r=>!isBlocked(r.email)),accent:"#581c87",renderCard:(r,i,ac)=><PromoCard key={i} r={r} accent={ac} badgeLabel="FAKE DOMAIN" onBlock={e=>openBlockModal(e,"promo")} isBlocked={isBlocked} blockingEmail={blockingEmail}/>},
    ],
    onReimport:()=>{resetPromo();setPromoModal(true);setPromoStep("drop");setPromoRows([]);setPromoErr("");},
    showRaw:promoShowRaw,setShowRaw:setPromoShowRaw,
    previewCols:["User Email","User Name","Promo Code","Discount Amount","Order Status","Card Number","Wallet Num","Created At"],
    previewKeyFn:(row,col)=>getCol(row)(col)||"—",
    onDownload:()=>{const wb=XLSX.utils.book_new();[["HighDiscount",promoHighDiscount],["SameCard",promoSameCard],["SameWallet",promoSameWallet],["FakeDomain",promoFakeDomain]].forEach(([name,rows])=>{const data=rows.map(r=>({"Email":r.email||r.card||r.wallet||"—","Names":(r.custNames||[]).join(", ")||"—","Promo Codes":(r.promoCodes||[]).join(", ")||"—","Discount (EGP)":r.totalDiscount||0,"Tx Count":r.txCount,"Reason":(r.reasons||[]).join(" | ")}));const ws=XLSX.utils.json_to_sheet(data.length>0?data:[{}]);XLSX.utils.book_append_sheet(wb,ws,name);});XLSX.writeFile(wb,`Promo_Abuse_${new Date().toISOString().split("T")[0]}.xlsx`);},
    ...blockProps,
  };

  // ── Empty state component ──
  const EmptyState=({icon,title,desc,accent,onImport,btnLabel})=>(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",padding:"80px 0",textAlign:"center"}}>
      <div style={{width:80,height:80,borderRadius:"50%",background:accent+"18",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:20}}><Shield size={36} color={accent}/></div>
      <div style={{fontWeight:800,fontSize:22,color:"#1e293b",marginBottom:10}}>{title}</div>
      <div style={{color:"#64748b",fontSize:14,lineHeight:1.8,marginBottom:20,maxWidth:480}} dangerouslySetInnerHTML={{__html:desc}}/>
      <button onClick={onImport} style={{display:"flex",alignItems:"center",gap:8,padding:"13px 30px",background:accent,color:"#fff",border:"none",borderRadius:10,cursor:"pointer",fontSize:14,fontWeight:700,boxShadow:`0 4px 14px ${accent}44`}}><Upload size={16}/> {btnLabel}</button>
    </div>
  );

  return(<div style={{minHeight:"100vh",background:"#f0f2f8",fontFamily:"'Inter',sans-serif",fontSize:14}}>
    {showAdminPanel&&<AdminPanel onClose={()=>setShowAdminPanel(false)}/>}
    {showChangePw&&<ChangeMyPasswordModal currentUser={currentUser} onClose={()=>setShowChangePw(false)}/>}

    {/* Header */}
    <div style={{background:"#0f172a",padding:"12px 28px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
      <div style={{display:"flex",flexDirection:"column",gap:4}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}><Shield size={24} color="#f97316"/><span style={{fontWeight:800,fontSize:16,color:"#f8fafc"}}>Waffarha Guard</span><span style={{fontSize:10,color:"#475569",background:"#1e293b",padding:"2px 9px",borderRadius:20}}>Fraud Control</span></div>
        <span style={{fontWeight:800,fontSize:12,color:"#f97316",letterSpacing:0.5}}>Waffarha</span>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{textAlign:"right"}}><div style={{fontSize:13,color:"#f8fafc",fontWeight:600}}>{currentUser.username}</div><div style={{fontSize:11,color:"#475569"}}>{currentUser.role==="superadmin"?"Super Admin":"User"}</div></div>
        {currentUser.role==="user"&&<button onClick={()=>setShowChangePw(true)} style={{display:"flex",alignItems:"center",gap:6,padding:"7px 14px",background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:600}}>🔑 Change Password</button>}
        {currentUser.role==="superadmin"&&<button onClick={()=>setShowAdminPanel(true)} style={{display:"flex",alignItems:"center",gap:6,padding:"7px 14px",background:"#f97316",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700}}><Users size={13}/> Manage Users</button>}
        <button onClick={()=>setCurrentUser(null)} style={{display:"flex",alignItems:"center",gap:6,padding:"7px 14px",background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:600}}>Sign out</button>
      </div>
    </div>

    {/* Sidebar toggle */}
    <button onClick={()=>setSidebarOpen(v=>!v)} style={{position:"fixed",top:72,left:sidebarOpen?236:8,zIndex:900,background:"#1e293b",border:"1px solid #334155",borderRadius:8,width:28,height:28,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",transition:"left 0.25s ease",boxShadow:"0 2px 8px rgba(0,0,0,0.3)"}}><span style={{color:"#94a3b8",fontSize:14,lineHeight:1}}>{sidebarOpen?"◀":"▶"}</span></button>

    {/* Sidebar */}
    <div style={{position:"fixed",top:62,left:0,bottom:0,width:sidebarOpen?248:0,background:"#0f172a",borderRight:"1px solid #1e293b",zIndex:800,overflow:"hidden",transition:"width 0.25s ease",display:"flex",flexDirection:"column"}}>
      <div style={{width:248,padding:"18px 0 12px",flex:1,overflowY:"auto"}}>
        <div style={{padding:"0 16px 12px",fontSize:10,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:1}}>Platforms</div>
        {PLATFORM_TABS.filter(pt=>!pt.adminOnly||(pt.adminOnly&&currentUser.role==="superadmin")).map(pt=>{const badges={paytabs:ptBadge,noon:noonBadge,paymob:pmBadge,admin:adminBadge,fawry:fawryBadge,promo:promoBadge};const count=badges[pt.id]||0;const isActive=tab===pt.id;return(<button key={pt.id} onClick={()=>pt.ready&&setTab(pt.id)} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"11px 16px",border:"none",cursor:pt.ready?"pointer":"default",background:isActive?`${pt.clr}22`:"transparent",borderLeft:isActive?`3px solid ${pt.clr}`:"3px solid transparent",textAlign:"left",transition:"background 0.15s"}}><div style={{width:8,height:8,borderRadius:"50%",background:pt.ready?pt.clr:"#334155",flexShrink:0}}/><span style={{fontSize:13,fontWeight:isActive?700:500,color:isActive?pt.clr:"#94a3b8",flex:1}}>{pt.label}</span>{!pt.ready&&<span style={{fontSize:9,color:"#334155",background:"#1e293b",padding:"1px 6px",borderRadius:8}}>SOON</span>}{pt.ready&&count>0&&<span style={{background:"#dc2626",color:"#fff",fontSize:10,padding:"1px 7px",borderRadius:10,fontWeight:700,flexShrink:0}}>{count}</span>}</button>);})}
      </div>
    </div>

    {/* Main */}
    <div style={{marginLeft:sidebarOpen?248:0,transition:"margin-left 0.25s ease",paddingTop:8}}>

      {/* PayTabs */}
      {tab==="paytabs"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {ptRaw.length===0?<EmptyState accent="#7c3aed" title="PayTabs Fraud Detection" desc='Auto-creates <b style="color:#7c3aed">CC Details</b> = Payment Description + Expiry Year/Month' onImport={()=>{setPtModal(true);setPtStep("drop");setPtRows([]);setPtErr("");}} btnLabel="Import PayTabs Sheet"/>:<UniversalDashboard config={ptConfig}/>}
      </div>}

      {/* Noon */}
      {tab==="noon"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {noonMerged.length===0?(<>
          <div style={{fontWeight:800,fontSize:22,color:"#1e293b",marginBottom:6}}>Noon Payment Detection</div>
          <div style={{color:"#64748b",fontSize:13,lineHeight:1.7,marginBottom:24}}>Upload both files to merge and analyze. <b style={{color:"#f59e0b"}}>CC Details</b> = <code>Payerinfo</code></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:18}}>
            {[{key:"noon",label:"NOON File",desc:"Payment transaction data",ref:noonRef,state:noonRaw,setState:setNoonRaw,drag:noonDrag,setDrag:setNoonDrag},{key:"admin",label:"ADMIN File",desc:"User & order admin data",ref:adminNoonRef,state:adminRawFile,setState:setAdminRawFile,drag:adminDrag,setDrag:setAdminDrag}].map(f=>(
              <div key={f.key} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
                <div style={{fontWeight:700,color:"#0f172a",marginBottom:3,fontSize:15}}>📄 {f.label}</div>
                <div style={{fontSize:12,color:"#64748b",marginBottom:14}}>{f.desc}</div>
                {f.state?(<div style={{background:"#f0fdf4",borderRadius:10,padding:"12px 16px",display:"flex",alignItems:"center",gap:12}}><CheckCircle size={18} color="#16a34a"/><div style={{flex:1}}><div style={{fontWeight:600,color:"#0f172a",fontSize:13}}>{f.state.file.name}</div><div style={{fontSize:11,color:"#16a34a",marginTop:2}}>{f.state.rows.length.toLocaleString()} rows</div></div><button onClick={()=>f.setState(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8"}}><X size={15}/></button></div>)
                :(<div onClick={()=>f.ref.current.click()} onDragOver={e=>{e.preventDefault();f.setDrag(true);}} onDragLeave={()=>f.setDrag(false)} onDrop={e=>{e.preventDefault();f.setDrag(false);if(e.dataTransfer.files[0])loadNoonFile(e.dataTransfer.files[0],f.key);}} style={{border:`2px dashed ${f.drag?"#f59e0b":"#e2e8f0"}`,borderRadius:12,padding:36,textAlign:"center",cursor:"pointer",background:f.drag?"#fffbeb":"#fafafa"}}><Upload size={22} color={f.drag?"#f59e0b":"#94a3b8"} style={{margin:"0 auto 10px",display:"block"}}/><div style={{fontWeight:600,color:"#334155",fontSize:13,marginBottom:4}}>Drop {f.label} here</div><div style={{fontSize:12,color:"#94a3b8"}}>CSV · Excel</div><input ref={f.ref} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])loadNoonFile(e.target.files[0],f.key);e.target.value="";}} /></div>)}
              </div>
            ))}
          </div>
          <div style={{textAlign:"center"}}><button onClick={doNoonAnalyze} disabled={!noonRaw||!adminRawFile||noonLoading} style={{display:"inline-flex",alignItems:"center",gap:10,padding:"13px 36px",background:(!noonRaw||!adminRawFile)?"#e2e8f0":"#f59e0b",color:(!noonRaw||!adminRawFile)?"#94a3b8":"#fff",border:"none",borderRadius:10,cursor:(!noonRaw||!adminRawFile)?"not-allowed":"pointer",fontSize:14,fontWeight:800}}>🔀 {noonLoading?"Merging…":"Merge & Analyse"}</button></div>
        </>):<UniversalDashboard config={noonConfig}/>}
      </div>}

      {/* PayMob */}
      {tab==="paymob"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {pmRaw.length===0?<EmptyState accent="#2563eb" title="PayMob Detection" desc='Detects fraud across <b style="color:#2563eb">Digital Wallets</b> and <b style="color:#7c3aed">BNPL</b> transactions.' onImport={()=>{setPmModal(true);setPmStep("drop");setPmRows([]);setPmErr("");}} btnLabel="Import PayMob Sheet"/>:(<>
          <div style={{display:"flex",gap:0,marginBottom:20,background:"#fff",borderRadius:12,padding:6,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",width:"fit-content"}}>
            {[{id:"wallets",label:"📱 Wallets",count:pmWFraud.length+pmWHighAmt.length+pmWFakeDom.length+pmWalletAbusers.length},{id:"bnpl",label:"💳 BNPL",count:pmBFraud.length+pmBHighAmt.length}].map(st=>(<button key={st.id} onClick={()=>setPmSubTab(st.id)} style={{padding:"9px 24px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:pmSubTab===st.id?"#2563eb":"transparent",color:pmSubTab===st.id?"#fff":"#64748b",display:"flex",alignItems:"center",gap:8}}>{st.label}{st.count>0&&<span style={{background:pmSubTab===st.id?"rgba(255,255,255,0.3)":"#e2e8f0",color:pmSubTab===st.id?"#fff":"#64748b",fontSize:11,padding:"1px 7px",borderRadius:10,fontWeight:700}}>{st.count}</span>}</button>))}
          </div>
          {pmSubTab==="wallets"&&<UniversalDashboard config={pmWConfig}/>}
          {pmSubTab==="bnpl"&&<UniversalDashboard config={pmBConfig}/>}
        </>)}
      </div>}

      {/* Admin */}
      {tab==="admin"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {adminSheetRaw.length===0?<EmptyState accent="#0ea5e9" title="Admin Transaction Detection" desc='Detects payment method abuse, biller abuse, high amounts, fake domains &amp; recharge abusers.' onImport={()=>{setAdminModal(true);setAdminStep("drop");setAdminRows([]);setAdminErr("");}} btnLabel="Import Admin Sheet"/>:<UniversalDashboard config={adminConfig}/>}
      </div>}

      {/* Fawry */}
      {tab==="fawry"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {fawryRaw.length===0?<EmptyState accent="#f97316" title="Fawry Fraud Detection" desc='Pre-processes the <b style="color:#f97316">user</b> column, then detects suspected, high amounts &amp; fake domains.' onImport={()=>{setFawryModal(true);setFawryStep("drop");setFawryRows([]);setFawryErr("");}} btnLabel="Import Fawry Sheet"/>:<UniversalDashboard config={fawryConfig}/>}
      </div>}

      {/* Promo */}
      {tab==="promo"&&<div style={{padding:"24px 28px",maxWidth:1500,margin:"0 auto"}}>
        {promoRaw.length===0?<EmptyState accent="#e11d48" title="Promo Code Abuse Detection" desc='Detects high discount abuse, shared card/wallet across accounts &amp; fake domain promo claims.' onImport={()=>{setPromoModal(true);setPromoStep("drop");setPromoRows([]);setPromoErr("");}} btnLabel="Import Promo Sheet"/>:<UniversalDashboard config={promoConfig}/>}
      </div>}

      {/* Audit Log */}
      {tab==="dashboard"&&currentUser.role==="superadmin"&&(()=>{
        const PLAT_META={paytabs:{label:"PayTabs",clr:"#7c3aed"},noon:{label:"Noon",clr:"#f59e0b"},paymob_wallets:{label:"PayMob W",clr:"#2563eb"},paymob_bnpl:{label:"BNPL",clr:"#6d28d9"},admin:{label:"Admin",clr:"#0ea5e9"},fawry:{label:"Fawry",clr:"#f97316"},promo:{label:"Promo",clr:"#e11d48"}};
        const TYPE_META={multi_cc:{label:"Multi CC",clr:"#ef4444"},high_amount:{label:"High Amount",clr:"#f97316"},fake_domain:{label:"Fake Domain",clr:"#8b5cf6"},wallet_abuser:{label:"Wallet Abuser",clr:"#06b6d4"},bnpl_fraud:{label:"BNPL Fraud",clr:"#ec4899"},pay_method_abuse:{label:"Pay Method",clr:"#10b981"},suspected_trials:{label:"Suspected",clr:"#f59e0b"},recharge_abuser:{label:"Recharge",clr:"#6366f1"},fawry_suspected:{label:"Fawry Susp.",clr:"#f97316"},promo_high_discount:{label:"Promo Disc.",clr:"#e11d48"},promo_same_card:{label:"Promo Card",clr:"#dc2626"},promo_same_wallet:{label:"Promo Wallet",clr:"#b91c1c"},promo_fake_domain:{label:"Promo FakeDom",clr:"#991b1b"}};

        // ── Computed ────────────────────────────────────────────────────────────
        const uniqueEmails=new Set(dashAlerts.map(a=>a.entity_email).filter(Boolean));
        const totalAmt=dashAlerts.reduce((s,a)=>s+(parseFloat(a.total_amount)||0),0);

        const platCounts={};
        dashAlerts.forEach(a=>{platCounts[a.platform]=(platCounts[a.platform]||0)+1;});
        const platArr=Object.entries(platCounts).map(([k,v])=>({key:k,label:PLAT_META[k]?.label||k,clr:PLAT_META[k]?.clr||"#94a3b8",value:v})).sort((a,b)=>b.value-a.value);
        const maxPlat=Math.max(...platArr.map(p=>p.value),1);

        const typeCounts={};
        dashAlerts.forEach(a=>{typeCounts[a.alert_type]=(typeCounts[a.alert_type]||0)+1;});
        const typeArr=Object.entries(typeCounts).map(([k,v])=>({key:k,label:TYPE_META[k]?.label||k,clr:TYPE_META[k]?.clr||"#94a3b8",value:v})).sort((a,b)=>b.value-a.value);
        const totalTypes=typeArr.reduce((s,t)=>s+t.value,0)||1;

        const DONUT_R=56,DONUT_CIRC=2*Math.PI*DONUT_R;
        let donutOffset=0;
        const donutSegs=typeArr.map(t=>{const pct=t.value/totalTypes;const dash=pct*DONUT_CIRC;const seg={...t,dash,gap:DONUT_CIRC-dash,offset:donutOffset};donutOffset+=dash;return seg;});

        const byDate={};
        dashAlerts.forEach(a=>{const d=(a.created_at||"").split("T")[0]||"unknown";if(d!=="unknown")byDate[d]=(byDate[d]||0)+1;});
        const sortedDates=Object.keys(byDate).sort();
        const displayDates=sortedDates.length>60?sortedDates.slice(-60):sortedDates;
        const maxDay=Math.max(...Object.values(byDate),1);

        const emailMap={};
        dashAlerts.filter(a=>a.entity_email).forEach(a=>{
          if(!emailMap[a.entity_email])emailMap[a.entity_email]={platforms:new Set(),count:0,amt:0,types:new Set()};
          emailMap[a.entity_email].platforms.add(a.platform);
          emailMap[a.entity_email].count++;
          emailMap[a.entity_email].amt+=(parseFloat(a.total_amount)||0);
          emailMap[a.entity_email].types.add(a.alert_type);
        });
        const repeatOffenders=Object.entries(emailMap)
          .filter(([,v])=>v.platforms.size>=2)
          .map(([email,v])=>({email,platforms:[...v.platforms],count:v.count,amt:v.amt,types:[...v.types],blocked:blockedEmails.has(email)}))
          .sort((a,b)=>b.platforms.length-a.platforms.length||b.count-a.count)
          .slice(0,25);

        const downloadReport=()=>{
          const wb=XLSX.utils.book_new();
          const summaryRows=Object.entries(emailMap).map(([email,v])=>[email,[...v.platforms].join(", "),v.count,Math.round(v.amt),[...v.types].join(", "),blockedEmails.has(email)?"Yes":"No"]).sort((a,b)=>b[2]-a[2]);
          XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Email","Platforms","Alert Count","Total EGP","Alert Types","Blocked"],...summaryRows]),"All Flagged Emails");
          const repeatRows=repeatOffenders.map(r=>[r.email,r.platforms.length,r.platforms.join(", "),r.count,Math.round(r.amt),r.blocked?"Yes":"No"]);
          XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Email","# Platforms","Platforms","Alert Count","Total EGP","Blocked"],...repeatRows]),"Cross-Platform Offenders");
          XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Platform","Alert Count"],...platArr.map(p=>[p.label,p.value])]),"By Platform");
          XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Date","Platform","Uploaded By","Records"],...[...dashSessions].reverse().map(s=>[new Date(s.created_at).toLocaleString(),s.platform,s.uploaded_by,s.record_count])]),"Import Sessions");
          XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Date","Platform","Alert Type","Risk","Email","Amount EGP","Tx Count"],...dashAlerts.map(a=>[new Date(a.created_at).toLocaleString(),a.platform,a.alert_type,a.risk_level,a.entity_email||"",parseFloat(a.total_amount)||0,a.transaction_count||0])]),"All Alerts");
          XLSX.writeFile(wb,`waffarha-report-${dashRange||"all"}d-${new Date().toISOString().split("T")[0]}.xlsx`);
        };

        return(
        <div style={{padding:"24px 28px",maxWidth:1300,margin:"0 auto"}}>
          {/* Header */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24,flexWrap:"wrap",gap:12}}>
            <div>
              <div style={{fontWeight:800,fontSize:20,color:"#0f172a"}}>📊 Historical Dashboard</div>
              <div style={{fontSize:13,color:"#64748b",marginTop:3}}>Accumulated fraud data from all import sessions · Supabase powered</div>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              {[{label:"7d",v:7},{label:"30d",v:30},{label:"90d",v:90},{label:"All time",v:null}].map(r=>(
                <button key={r.label} onClick={()=>setDashRange(r.v)} style={{padding:"7px 14px",background:dashRange===r.v?"#6366f1":"#f1f5f9",color:dashRange===r.v?"#fff":"#475569",border:"none",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700}}>{r.label}</button>
              ))}
              <button onClick={downloadReport} disabled={dashAlerts.length===0} style={{padding:"7px 16px",background:"#16a34a",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700,display:"flex",alignItems:"center",gap:6,opacity:dashAlerts.length===0?0.5:1}}>⬇ Download Excel</button>
              <button onClick={()=>loadDash(dashRange)} disabled={dashLoading} style={{padding:"7px 14px",background:"#f1f5f9",color:"#475569",border:"none",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700}}>{dashLoading?"…":"↻ Refresh"}</button>
            </div>
          </div>

          {dashLoading&&<div style={{textAlign:"center",padding:72,color:"#6366f1",fontSize:15,fontWeight:700}}>Loading historical data…</div>}
          {dashErr&&<div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:12,padding:"14px 18px",color:"#dc2626",fontSize:13,marginBottom:20}}>⚠ {dashErr}</div>}

          {!dashLoading&&dashAlerts.length===0&&(
            <div style={{background:"#fff",borderRadius:16,padding:72,textAlign:"center",color:"#94a3b8",fontSize:14,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontSize:44,marginBottom:14}}>📊</div>
              <div style={{fontWeight:700,fontSize:16,color:"#475569",marginBottom:8}}>No historical data yet</div>
              <div style={{marginBottom:12}}>Import data from any platform — each session is saved automatically to Supabase.</div>
              <div style={{fontSize:12,color:"#d97706",background:"#fef3c7",padding:"8px 16px",borderRadius:8,display:"inline-block",fontWeight:600}}>⚠ If you've imported data but see nothing, run in Supabase SQL Editor:<br/>ALTER TABLE fraud_alerts DISABLE ROW LEVEL SECURITY;<br/>ALTER TABLE upload_sessions DISABLE ROW LEVEL SECURITY;</div>
            </div>
          )}

          {!dashLoading&&dashAlerts.length>0&&<>
            {/* Stat cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:14,marginBottom:20}}>
              {[
                {icon:"📧",label:"Unique Emails",value:uniqueEmails.size.toLocaleString(),clr:"#6366f1"},
                {icon:"🚨",label:"Total Alerts",value:dashAlerts.length.toLocaleString(),clr:"#dc2626"},
                {icon:"💰",label:"EGP at Risk",value:`${Math.round(totalAmt).toLocaleString()}`,clr:"#f97316"},
                {icon:"📅",label:"Sessions",value:dashSessions.length.toLocaleString(),clr:"#10b981"},
                {icon:"🚫",label:"Blocked",value:blockedList.length.toLocaleString(),clr:"#ef4444"},
              ].map(s=>(
                <div key={s.label} style={{background:"#fff",borderRadius:14,padding:"18px 20px",boxShadow:"0 2px 8px rgba(0,0,0,0.07)",border:`1.5px solid ${s.clr}25`}}>
                  <div style={{fontSize:26,marginBottom:8}}>{s.icon}</div>
                  <div style={{fontWeight:900,fontSize:22,color:s.clr,lineHeight:1}}>{s.value}</div>
                  <div style={{fontSize:11,color:"#94a3b8",fontWeight:600,marginTop:5}}>{s.label}</div>
                </div>
              ))}
            </div>

            {/* Platform bar + Alert type donut */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
              <div style={{background:"#fff",borderRadius:16,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.07)"}}>
                <div style={{fontWeight:800,fontSize:15,color:"#0f172a",marginBottom:18}}>Alerts by Platform</div>
                <div style={{display:"flex",flexDirection:"column",gap:12}}>
                  {platArr.map(p=>(
                    <div key={p.key}>
                      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5}}>
                        <span style={{fontWeight:700,color:"#334155"}}>{p.label}</span>
                        <span style={{fontWeight:800,color:p.clr}}>{p.value}</span>
                      </div>
                      <div style={{height:10,background:"#f1f5f9",borderRadius:6,overflow:"hidden"}}>
                        <div style={{height:"100%",width:`${(p.value/maxPlat)*100}%`,background:p.clr,borderRadius:6}}/>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div style={{background:"#fff",borderRadius:16,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.07)"}}>
                <div style={{fontWeight:800,fontSize:15,color:"#0f172a",marginBottom:18}}>Alert Type Breakdown</div>
                <div style={{display:"flex",alignItems:"center",gap:20}}>
                  <div style={{flexShrink:0}}>
                    <svg width={140} height={140} viewBox="0 0 140 140">
                      {donutSegs.map((seg,i)=>(
                        <circle key={i} cx={70} cy={70} r={DONUT_R} fill="none" stroke={seg.clr} strokeWidth={24}
                          strokeDasharray={`${seg.dash} ${seg.gap}`} strokeDashoffset={-(seg.offset)}
                          transform="rotate(-90 70 70)" style={{opacity:0.9}}/>
                      ))}
                      <text x={70} y={65} textAnchor="middle" style={{fontSize:18,fontWeight:"900",fill:"#0f172a"}}>{dashAlerts.length}</text>
                      <text x={70} y={82} textAnchor="middle" style={{fontSize:10,fill:"#94a3b8"}}>total alerts</text>
                    </svg>
                  </div>
                  <div style={{flex:1,display:"flex",flexDirection:"column",gap:7,overflowY:"auto",maxHeight:160}}>
                    {typeArr.slice(0,9).map(t=>(
                      <div key={t.key} style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:10,height:10,borderRadius:"50%",background:t.clr,flexShrink:0}}/>
                        <span style={{fontSize:11,color:"#475569",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.label}</span>
                        <span style={{fontSize:12,fontWeight:700,color:t.clr}}>{t.value}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            {/* Daily trend */}
            <div style={{background:"#fff",borderRadius:16,padding:"20px 22px",boxShadow:"0 2px 8px rgba(0,0,0,0.07)",marginBottom:16}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                <div style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>Daily Alert Trend</div>
                <div style={{fontSize:12,color:"#94a3b8"}}>{displayDates.length} days</div>
              </div>
              {displayDates.length===0?(
                <div style={{color:"#94a3b8",fontSize:13,padding:"20px 0",textAlign:"center"}}>No data</div>
              ):(
                <div>
                  <div style={{display:"flex",alignItems:"flex-end",gap:displayDates.length>30?1:3,height:130}}>
                    {displayDates.map(d=>(
                      <div key={d} title={`${d}: ${byDate[d]} alerts`} style={{flex:1,display:"flex",flexDirection:"column",justifyContent:"flex-end",alignItems:"stretch"}}>
                        <div style={{background:"#6366f1",borderRadius:"2px 2px 0 0",minHeight:byDate[d]>0?3:0,height:`${(byDate[d]/maxDay)*100}%`,opacity:0.8}}/>
                      </div>
                    ))}
                  </div>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:8,fontSize:10,color:"#94a3b8",fontFamily:"monospace"}}>
                    <span>{displayDates[0]}</span>
                    {displayDates.length>4&&<span>{displayDates[Math.floor(displayDates.length/2)]}</span>}
                    <span>{displayDates[displayDates.length-1]}</span>
                  </div>
                </div>
              )}
            </div>

            {/* Cross-platform repeat offenders */}
            <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 8px rgba(0,0,0,0.07)",overflow:"hidden",marginBottom:16}}>
              <div style={{padding:"18px 22px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>Cross-Platform Repeat Offenders</div>
                  <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>Emails flagged on 2+ platforms — highest priority threats</div>
                </div>
                <div style={{background:"#fef2f2",color:"#dc2626",padding:"5px 14px",borderRadius:8,fontSize:13,fontWeight:700,flexShrink:0}}>{repeatOffenders.length} found</div>
              </div>
              {repeatOffenders.length===0?(
                <div style={{padding:40,textAlign:"center",color:"#94a3b8",fontSize:13}}>No cross-platform offenders in this period.</div>
              ):(
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr style={{background:"#1e293b"}}>{["Email","Platforms","Alerts","EGP at Risk","Status"].map(h=><th key={h} style={{padding:"11px 18px",textAlign:"left",fontSize:11,fontWeight:700,color:"#94a3b8",textTransform:"uppercase",letterSpacing:0.8,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                  <tbody>{repeatOffenders.map((r,i)=>(
                    <tr key={r.email} style={{borderBottom:"1px solid #f1f5f9",background:i%2===0?"#fafafa":"#fff"}}>
                      <td style={{padding:"12px 18px",fontWeight:700,fontSize:13,color:"#0f172a",wordBreak:"break-all"}}>{r.email}</td>
                      <td style={{padding:"12px 18px"}}>
                        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                          {r.platforms.map(p=><span key={p} style={{background:(PLAT_META[p]?.clr||"#94a3b8")+"20",color:PLAT_META[p]?.clr||"#94a3b8",padding:"2px 8px",borderRadius:5,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{PLAT_META[p]?.label||p}</span>)}
                        </div>
                      </td>
                      <td style={{padding:"12px 18px",fontWeight:800,color:"#dc2626",fontSize:14}}>{r.count}</td>
                      <td style={{padding:"12px 18px",fontWeight:700,color:"#f97316"}}>EGP {Math.round(r.amt).toLocaleString()}</td>
                      <td style={{padding:"12px 18px"}}>
                        {r.blocked
                          ?<span style={{background:"#fee2e2",color:"#dc2626",padding:"3px 10px",borderRadius:6,fontSize:12,fontWeight:700}}>🚫 Blocked</span>
                          :<button onClick={()=>openBlockModal(r.email,"dashboard")} style={{background:"#fef2f2",border:"1px solid #fca5a5",color:"#dc2626",padding:"3px 10px",borderRadius:6,fontSize:12,fontWeight:700,cursor:"pointer"}}>Block</button>}
                      </td>
                    </tr>
                  ))}</tbody>
                </table>
              )}
            </div>

            {/* Import session history */}
            <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 8px rgba(0,0,0,0.07)",overflow:"hidden"}}>
              <div style={{padding:"18px 22px",borderBottom:"1px solid #f1f5f9"}}>
                <div style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>Import Session History</div>
                <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>{dashSessions.length} sessions in selected period</div>
              </div>
              {dashSessions.length===0?(
                <div style={{padding:40,textAlign:"center",color:"#94a3b8",fontSize:13}}>No sessions in this period.</div>
              ):(
                <div style={{maxHeight:320,overflowY:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr style={{background:"#1e293b",position:"sticky",top:0}}>{["Date & Time","Platform","Uploaded By","Records"].map(h=><th key={h} style={{padding:"11px 18px",textAlign:"left",fontSize:11,fontWeight:700,color:"#94a3b8",textTransform:"uppercase",letterSpacing:0.8,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                    <tbody>{[...dashSessions].reverse().map((s,i)=>(
                      <tr key={s.id} style={{borderBottom:"1px solid #f1f5f9",background:i%2===0?"#fafafa":"#fff"}}>
                        <td style={{padding:"11px 18px",fontSize:12,fontFamily:"monospace",color:"#64748b",whiteSpace:"nowrap"}}>{new Date(s.created_at).toLocaleString()}</td>
                        <td style={{padding:"11px 18px"}}><span style={{background:(PLAT_META[s.platform]?.clr||"#94a3b8")+"20",color:PLAT_META[s.platform]?.clr||"#94a3b8",padding:"3px 10px",borderRadius:6,fontSize:12,fontWeight:700}}>{PLAT_META[s.platform]?.label||s.platform}</span></td>
                        <td style={{padding:"11px 18px",fontSize:12,fontWeight:600,color:"#334155"}}>{s.uploaded_by}</td>
                        <td style={{padding:"11px 18px",fontSize:13,fontWeight:700,color:"#0f172a"}}>{(s.record_count||0).toLocaleString()}</td>
                      </tr>
                    ))}</tbody>
                  </table>
                </div>
              )}
            </div>
          </>}
        </div>);
      })()}

      {tab==="audit"&&currentUser.role==="superadmin"&&(()=>{
        const logs=getAuditLog();
        const downloadAudit=()=>{const wb=XLSX.utils.book_new();const data=logs.map(l=>({"Time":l.time,"User":l.user,"Action":l.action,"Platform":l.platform,"Records":l.records||"","Details":l.details||""}));const ws=XLSX.utils.json_to_sheet(data.length>0?data:[{}]);ws["!cols"]=[22,18,12,14,10,60].map(wch=>({wch}));XLSX.utils.book_append_sheet(wb,ws,"Audit Log");XLSX.writeFile(wb,`Audit_Log_${new Date().toISOString().split("T")[0]}.xlsx`);};
        const platformClr={PayTabs:"#7c3aed",Noon:"#f59e0b",PayMob:"#2563eb",Admin:"#0ea5e9",Fawry:"#f97316"};
        return(<div style={{padding:"24px 28px",maxWidth:1400,margin:"0 auto"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}><div><div style={{fontWeight:800,fontSize:18,color:"#0f172a"}}>Audit Log</div><div style={{fontSize:13,color:"#64748b",marginTop:3}}>{logs.length} event{logs.length!==1?"s":""} this session</div></div><button onClick={downloadAudit} disabled={logs.length===0} style={{display:"flex",alignItems:"center",gap:6,padding:"9px 18px",background:logs.length===0?"#e2e8f0":"#10b981",color:logs.length===0?"#94a3b8":"#fff",border:"none",borderRadius:8,cursor:logs.length===0?"not-allowed":"pointer",fontSize:13,fontWeight:700}}><Download size={14}/> Export .xlsx</button></div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:24}}>{[{label:"Total Events",value:logs.length,clr:"#10b981"},{label:"Imports",value:logs.filter(l=>l.action==="Import").length,clr:"#2563eb"},{label:"Platforms",value:new Set(logs.map(l=>l.platform).filter(Boolean)).size,clr:"#f97316"}].map(({label,value,clr})=>(<div key={label} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderLeft:`5px solid ${clr}`}}><div style={{fontSize:11,color:"#64748b",fontWeight:600,textTransform:"uppercase",letterSpacing:0.5,marginBottom:6}}>{label}</div><div style={{fontSize:30,fontWeight:900,color:clr,lineHeight:1}}>{value}</div></div>))}</div>
          <div style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 8px rgba(0,0,0,0.08)",overflow:"hidden"}}>{logs.length===0?(<div style={{padding:60,textAlign:"center",color:"#94a3b8",fontSize:14}}>📋 No audit events yet.</div>):(<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr style={{background:"#1e293b"}}>{["#","Time","User","Action","Platform","Records","Details"].map(h=>(<th key={h} style={{padding:"12px 18px",textAlign:"left",fontWeight:700,color:"#94a3b8",fontSize:11,whiteSpace:"nowrap",letterSpacing:0.8,textTransform:"uppercase"}}>{h}</th>))}</tr></thead><tbody>{logs.map((l,i)=>(<tr key={l.id} style={{borderBottom:"1px solid #f1f5f9",background:i%2===0?"#fafafa":"#fff",verticalAlign:"top"}}><td style={{padding:"14px 18px",color:"#94a3b8",fontSize:12,fontWeight:600}}>#{logs.length-i}</td><td style={{padding:"14px 18px",fontSize:12,fontFamily:"monospace",color:"#334155",whiteSpace:"nowrap"}}>{l.time}</td><td style={{padding:"14px 18px"}}><div style={{display:"inline-flex",alignItems:"center",gap:6,background:"#1e293b",color:"#a5b4fc",padding:"4px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{l.user}</div></td><td style={{padding:"14px 18px"}}><div style={{display:"inline-flex",alignItems:"center",gap:5,background:"#dcfce7",color:"#16a34a",padding:"4px 10px",borderRadius:6,fontSize:12,fontWeight:700}}>📥 {l.action}</div></td><td style={{padding:"14px 18px"}}><div style={{display:"inline-flex",alignItems:"center",gap:6,padding:"4px 12px",borderRadius:8,fontSize:12,fontWeight:700,background:`${platformClr[l.platform]||"#64748b"}18`,color:platformClr[l.platform]||"#64748b"}}>{l.platform}</div></td><td style={{padding:"14px 18px",fontWeight:700,color:"#0f172a",fontSize:14}}>{(l.records||0).toLocaleString()}</td><td style={{padding:"14px 18px",fontSize:12,color:"#475569",maxWidth:400,lineHeight:1.7}}>{l.details||"—"}</td></tr>))}</tbody></table></div>)}</div>
        </div>);
      })()}

      {/* Settings */}
      {tab==="settings"&&currentUser.role==="superadmin"&&(()=>{
        // Build per-email activity across all loaded platforms
        const PLATFORM_SOURCES=[
          {name:"PayTabs",clr:"#7c3aed",enriched:ptEnriched,amtKey:"_amt"},
          {name:"Noon",clr:"#f59e0b",enriched:noonEnriched,amtKey:"_amt"},
          {name:"PayMob",clr:"#2563eb",enriched:pmWEnriched,amtKey:"_amt"},
          {name:"PayMob BNPL",clr:"#7c3aed",enriched:pmBEnriched,amtKey:"_amt"},
          {name:"Admin",clr:"#0ea5e9",enriched:adminEnriched,amtKey:"_amt"},
          {name:"Fawry",clr:"#f97316",enriched:fawryEnriched,amtKey:"_requested"},
          {name:"Promo",clr:"#e11d48",enriched:promoEnriched,amtKey:"_discountAmt"},
        ];
        const blockedActivity=blockedList.map(b=>{
          const key=b.entity_value.toLowerCase().trim();
          const platforms=PLATFORM_SOURCES.map(src=>{
            const rows=(src.enriched||[]).filter(r=>(r._email||"").toLowerCase()===key);
            if(rows.length===0)return null;
            const total=rows.reduce((s,r)=>s+(parseFloat(r[src.amtKey])||0),0);
            return{name:src.name,clr:src.clr,count:rows.length,total};
          }).filter(Boolean);
          const totalTx=platforms.reduce((s,p)=>s+p.count,0);
          const totalAmt=platforms.reduce((s,p)=>s+p.total,0);
          return{...b,platforms,totalTx,totalAmt};
        });
        const hasAnyData=PLATFORM_SOURCES.some(s=>(s.enriched||[]).length>0);
        return(
        <div style={{padding:"24px 28px",maxWidth:1200,margin:"0 auto"}}>
          {/* Header */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
            <div>
              <div style={{fontWeight:800,fontSize:20,color:"#0f172a"}}>🚫 Blocked Accounts</div>
              <div style={{fontSize:13,color:"#64748b",marginTop:3}}>Hidden globally across all platforms · activity shown from current session data</div>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"center"}}>
              {!hasAnyData&&<span style={{fontSize:12,color:"#d97706",background:"#fef3c7",padding:"5px 12px",borderRadius:8,fontWeight:600}}>⚠ Import data to see activity</span>}
              <div style={{fontSize:13,fontWeight:700,color:"#dc2626",background:"#fef2f2",border:"1px solid #fca5a5",padding:"7px 16px",borderRadius:9}}>{blockedList.length} blocked</div>
            </div>
          </div>

          {blockedList.length===0?(
            <div style={{background:"#fff",borderRadius:16,padding:72,textAlign:"center",color:"#94a3b8",fontSize:14,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontSize:40,marginBottom:12}}>✅</div>
              <div style={{fontWeight:700,fontSize:16,color:"#475569",marginBottom:6}}>No blocked accounts yet</div>
              <div>Use the <b>🚫 Block</b> button on any fraud card to add one.</div>
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              {blockedActivity.map((b,i)=>(
                <div key={b.entity_value} style={{background:"#fff",borderRadius:16,boxShadow:"0 2px 12px rgba(0,0,0,0.07)",border:"1.5px solid #fee2e2",overflow:"hidden"}}>
                  {/* Card header */}
                  <div style={{display:"flex",alignItems:"stretch"}}>
                    {/* Red block badge */}
                    <div style={{background:"#dc2626",width:72,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"16px 8px",gap:4}}>
                      <span style={{fontSize:22}}>🚫</span>
                      <div style={{color:"rgba(255,255,255,0.8)",fontSize:10,fontWeight:700,letterSpacing:0.5}}>BLOCKED</div>
                    </div>
                    {/* Email + meta */}
                    <div style={{flex:1,padding:"16px 20px",borderRight:"1px solid #f1f5f9"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,flexWrap:"wrap"}}>
                        <span style={{fontWeight:800,fontSize:15,color:"#0f172a",wordBreak:"break-all"}}>{b.entity_value}</span>
                        {isDisposable(b.entity_value)&&<span style={{background:"#fee2e2",color:"#b91c1c",padding:"2px 7px",borderRadius:4,fontSize:11,fontWeight:700,flexShrink:0}}>⚠ Non-wl</span>}
                      </div>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                        <span style={{background:"#1e293b",color:"#a5b4fc",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{b.blocked_by}</span>
                        {b.platform&&<span style={{background:"#f0f9ff",color:"#0369a1",padding:"3px 10px",borderRadius:6,fontSize:12,fontWeight:600}}>{b.platform}</span>}
                        <span style={{fontSize:12,color:"#94a3b8",fontFamily:"monospace"}}>{b.created_at?new Date(b.created_at).toLocaleString():"—"}</span>
                      </div>
                    </div>
                    {/* Reason */}
                    <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",justifyContent:"center",minWidth:220,maxWidth:320,borderRight:"1px solid #f1f5f9"}}>
                      <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:6}}>Reason</div>
                      <div style={{fontSize:13,color:"#334155",lineHeight:1.5,fontStyle:b.note?"normal":"italic"}}>{b.note||<span style={{color:"#cbd5e1"}}>No reason provided</span>}</div>
                    </div>
                    {/* Stats */}
                    <div style={{background:"#f8fafc",padding:"16px 20px",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:110,borderRight:"1px solid #f1f5f9"}}>
                      <div style={{fontSize:10,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Platforms</div>
                      <div style={{fontWeight:900,color:"#dc2626",fontSize:26,lineHeight:1}}>{b.platforms.length}</div>
                      <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>{b.totalTx} tx</div>
                    </div>
                    {/* Unblock */}
                    <div style={{padding:"16px 14px",display:"flex",alignItems:"center",justifyContent:"center"}}>
                      <button onClick={()=>handleUnblock(b.entity_value)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4,padding:"10px 14px",background:"#f0fdf4",border:"1.5px solid #86efac",borderRadius:10,cursor:"pointer"}}>
                        <span style={{fontSize:18}}>🔓</span>
                        <span style={{fontSize:10,fontWeight:700,color:"#16a34a",whiteSpace:"nowrap"}}>Unblock</span>
                      </button>
                    </div>
                  </div>
                  {/* Platform activity row */}
                  {b.platforms.length>0&&(
                    <div style={{borderTop:"1px solid #f1f5f9",padding:"12px 20px",display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                      <span style={{fontSize:11,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginRight:4}}>Activity:</span>
                      {b.platforms.map(p=>(
                        <div key={p.name} style={{display:"flex",alignItems:"center",gap:6,background:`${p.clr}12`,border:`1px solid ${p.clr}30`,borderRadius:8,padding:"5px 12px"}}>
                          <span style={{width:8,height:8,borderRadius:"50%",background:p.clr,display:"inline-block",flexShrink:0}}/>
                          <span style={{fontSize:12,fontWeight:700,color:p.clr}}>{p.name}</span>
                          <span style={{fontSize:12,color:"#64748b"}}>{p.count} tx</span>
                          {p.total>0&&<span style={{fontSize:12,color:"#0f172a",fontWeight:600}}>· EGP {p.total.toLocaleString()}</span>}
                        </div>
                      ))}
                    </div>
                  )}
                  {b.platforms.length===0&&hasAnyData&&(
                    <div style={{borderTop:"1px solid #f1f5f9",padding:"10px 20px"}}>
                      <span style={{fontSize:12,color:"#94a3b8"}}>No activity found in current session data for this email.</span>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>);
      })()}
    </div>

    {/* Modals */}
    <ImportModal open={ptModal} title="Import PayTabs Sheet" accent="#7c3aed" onClose={()=>setPtModal(false)} onImport={doImportPT} rows={ptRows} setRows={setPtRows} step={ptStep} setStep={setPtStep} previewCols={["Customer Email","Customer Name","Payment Description","Expiry Year","Expiry Month","Cart ID"]} loading={ptLoading} setLoading={setPtLoading} err={ptErr} setErr={setPtErr}/>
    <ImportModal open={pmModal} title="Import PayMob Sheet" accent="#2563eb" onClose={()=>setPmModal(false)} onImport={doImportPM} rows={pmRows} setRows={setPmRows} step={pmStep} setStep={setPmStep} previewCols={["client_email","client_name","client_phone","payment_method","amount_whole","success"]} loading={pmLoading} setLoading={setPmLoading} err={pmErr} setErr={setPmErr}/>
    <ImportModal open={adminModal} title="Import Admin Sheet" accent="#0ea5e9" onClose={()=>setAdminModal(false)} onImport={doImportAdmin} rows={adminRows} setRows={setAdminRows} step={adminStep} setStep={setAdminStep} previewCols={["User Email","User Name","Payment Method","Service Biller","Status","Requested Amount"]} loading={adminLoading} setLoading={setAdminLoading} err={adminErr} setErr={setAdminErr}/>
    <ImportModal open={fawryModal} title="Import Fawry Sheet" accent="#f97316" onClose={()=>setFawryModal(false)} onImport={doImportFawry} rows={fawryRows} setRows={setFawryRows} step={fawryStep} setStep={setFawryStep} previewCols={["user","user email","Requested","Fawry Code","Status","Payment Method"]} loading={fawryLoading} setLoading={setFawryLoading} err={fawryErr} setErr={setFawryErr}/>
    <ImportModal open={promoModal} title="Import Promo Sheet" accent="#e11d48" onClose={()=>setPromoModal(false)} onImport={doImportPromo} rows={promoRows} setRows={setPromoRows} step={promoStep} setStep={setPromoStep} previewCols={["User Email","User Name","Promo Code","Discount Amount","Order Status","Card Number","Wallet Num"]} loading={promoLoading} setLoading={setPromoLoading} err={promoErr} setErr={setPromoErr}/>

    {/* Block reason modal */}
    {blockModal&&<div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget){setBlockModal(null);setBlockNote("");}}}>
      <div style={{background:"#fff",borderRadius:18,padding:"32px 36px",width:420,boxShadow:"0 20px 60px rgba(0,0,0,0.18)",border:"1px solid #fee2e2"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20}}>
          <div style={{width:40,height:40,borderRadius:10,background:"#fef2f2",border:"1.5px solid #fca5a5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🚫</div>
          <div><div style={{fontWeight:800,fontSize:17,color:"#0f172a"}}>Block Account</div><div style={{fontSize:12,color:"#64748b",marginTop:2}}>This email will be hidden across all platforms</div></div>
        </div>
        <div style={{background:"#f8fafc",borderRadius:10,padding:"12px 14px",marginBottom:20,border:"1px solid #e2e8f0"}}>
          <div style={{fontSize:11,color:"#94a3b8",fontWeight:700,textTransform:"uppercase",letterSpacing:0.5,marginBottom:4}}>Email</div>
          <div style={{fontSize:14,fontWeight:700,color:"#dc2626",wordBreak:"break-all"}}>{blockModal.email}</div>
        </div>
        <div style={{marginBottom:24}}>
          <label style={{fontSize:13,fontWeight:700,color:"#334155",display:"block",marginBottom:8}}>Reason <span style={{color:"#dc2626"}}>*</span></label>
          <textarea value={blockNote} onChange={e=>{setBlockNote(e.target.value);if(e.target.value.trim())setBlockNoteErr(false);}} placeholder="e.g. Multiple chargeback requests, confirmed fraud…" rows={3} style={{width:"100%",boxSizing:"border-box",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${blockNoteErr?"#dc2626":"#e2e8f0"}`,fontSize:13,color:"#334155",outline:"none",resize:"vertical",fontFamily:"inherit",background:blockNoteErr?"#fef2f2":"#fff"}} onKeyDown={e=>{if(e.key==="Enter"&&e.ctrlKey)confirmBlock();}}/>
          {blockNoteErr&&<div style={{color:"#dc2626",fontSize:12,fontWeight:600,marginTop:4}}>⚠ Reason is required before blocking.</div>}
          <div style={{fontSize:11,color:"#94a3b8",marginTop:4}}>Ctrl+Enter to confirm</div>
        </div>
        <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
          <button onClick={()=>{setBlockModal(null);setBlockNote("");}} style={{padding:"9px 20px",background:"#f1f5f9",border:"none",borderRadius:9,cursor:"pointer",fontSize:13,fontWeight:600,color:"#64748b"}}>Cancel</button>
          <button onClick={confirmBlock} style={{padding:"9px 22px",background:"#dc2626",border:"none",borderRadius:9,cursor:"pointer",fontSize:13,fontWeight:700,color:"#fff",display:"flex",alignItems:"center",gap:6}}>🚫 Block</button>
        </div>
      </div>
    </div>}
  </div>);
}