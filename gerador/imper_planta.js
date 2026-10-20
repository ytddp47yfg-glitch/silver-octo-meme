// Pinta a planta de impermeabilização por período de execução (gerador/impermeabilizacao.py): classes pelas cores das
// hachuras do projeto (A vermelho, B magenta, C azul, D verde, ciano/amarelo fora do cronograma) + fechamento morfológico.
// Uso: node gerador/imper_planta.js planta.png planta_cor.jpg
const { chromium } = (() => { try { return require('playwright'); } catch (e) { return require(process.env.PW || '/opt/node22/lib/node_modules/playwright'); } })();const fs=require('fs');
(async()=>{const b=await chromium.launch();const p=await b.newPage();
 const data='data:image/png;base64,'+fs.readFileSync(process.argv[2]).toString('base64');
 const out=await p.evaluate(async ([d,cores])=>{const im=new Image();im.src=d;await im.decode();const W=im.width,H=im.height;
  const c=document.createElement('canvas');c.width=W;c.height=H;const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,W,H);x.drawImage(im,0,0);
  const id=x.getImageData(0,0,W,H),px=id.data;
  const cls=new Uint8Array(W*H);   // 1 A vermelho, 2 B magenta, 3 C azul, 4 D verde, 5 ciano, 6 amarelo
  for(let i=0,k=0;i<px.length;i+=4,k++){const r=px[i],g=px[i+1],bb=px[i+2],mx=Math.max(r,g,bb),mn=Math.min(r,g,bb);if(mx-mn<45)continue;
    if(bb>r+50&&bb>g+50)cls[k]=3;                                   // azul (piscinas, C)
    else if(r>170&&bb>170&&g<r-50&&Math.abs(r-bb)<50)cls[k]=2;      // magenta (B)
    else if(r>170&&g<r-50&&bb<r-50&&Math.abs(g-bb)<40)cls[k]=1;     // vermelho (A)
    else if(g>170&&bb>170&&r<g-50)cls[k]=5;                         // ciano (argamassa)
    else if(r>170&&g>170&&bb<g-50)cls[k]=6;                         // amarelo (resina/cristalizante)
    else if(g>r+50&&g>bb+50)cls[k]=4}                               // verde (D)
  const R=7;
  const morf=(m,dil)=>{const t=new Uint8Array(W*H),o=new Uint8Array(W*H);
    for(let y=0;y<H;y++)for(let xx=0;xx<W;xx++){let v=dil?0:1;for(let k=-R;k<=R&&(dil?!v:v);k++){const q=xx+k;const s=q<0||q>=W?(dil?0:1):m[y*W+q];v=dil?(v||s):(v&&s)}t[y*W+xx]=v}
    for(let y=0;y<H;y++)for(let xx=0;xx<W;xx++){let v=dil?0:1;for(let k=-R;k<=R&&(dil?!v:v);k++){const q=y+k;const s=q<0||q>=H?(dil?0:1):t[q*W+xx];v=dil?(v||s):(v&&s)}o[y*W+xx]=v}return o};
  const masks={};for(const k of [1,2,3,4,5,6]){let m=new Uint8Array(W*H);for(let i=0;i<W*H;i++)m[i]=cls[i]===k?1:0;masks[k]=morf(morf(m,1),0)}
  // tênis (B): faixa magenta dentro do retângulo da quadra (canto superior direito)
  const per=new Uint8Array(W*H);   // período por pixel
  for(let i=0;i<W*H;i++){const y=(i/W)|0,xx=i%W;
    if(masks[3][i])per[i]=3;else if(masks[1][i])per[i]=1;else if(masks[2][i])per[i]=(xx>755&&y<282)?5:2;else if(masks[4][i])per[i]=4;else if(masks[5][i]||masks[6][i])per[i]=6}
  // base: planta clareada em cinza; cor do período por cima; linhas escuras por cima de tudo
  for(let i=0,k=0;i<px.length;i+=4,k++){const r=px[i],g=px[i+1],bb=px[i+2],lum=0.3*r+0.59*g+0.11*bb,sat=Math.max(r,g,bb)-Math.min(r,g,bb);
    let v=sat>45?245:255-(255-lum)*0.55;let o=[v,v,v];
    if(per[k]){const cc=cores[per[k]],a=per[k]===6?0.35:0.82;o=o.map((t,j)=>t*(1-a)+cc[j]*a)}
    if(sat<45&&lum<110){o=[lum*0.6,lum*0.6,lum*0.6]}
    px[i]=o[0];px[i+1]=o[1];px[i+2]=o[2];px[i+3]=255}
  x.putImageData(id,0,0);return c.toDataURL('image/jpeg',0.86)},[data,{1:[33,102,172],2:[77,175,74],3:[255,170,0],4:[215,48,39],5:[118,42,131],6:[190,190,190]}]);
 fs.writeFileSync(process.argv[3],Buffer.from(out.split(',')[1],'base64'));await b.close()})();
