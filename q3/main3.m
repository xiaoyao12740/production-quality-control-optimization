%main3
%逐情形(A/B/C/D)期望成本+总体4096（利润降序）+绘图
%A: 零配件{1,2,3}→半成品1  B: 零配件{4,5,6}→半成品2
%C: 零配件{7,8}→半成品3    D: 半成品{1,2,3}→成品
%总体位：8(来料检)+1(半检)+1(半拆)+1(成检)+1(成拆)=12位，共4096
%返修口径：成检=否 且 半检=是→售后返修E=(cDF+cAF+cTF)/yF（含首轮拆）
clc;clear;close all;
%输出配置
OUT_ROOT=fullfile(pwd,'q3_output');if~exist(OUT_ROOT,'dir'),mkdir(OUT_ROOT);end
DIR_AB=fullfile(OUT_ROOT,'情形AB（共用）');if~exist(DIR_AB,'dir'),mkdir(DIR_AB);end
DIR_C=fullfile(OUT_ROOT,'情形C');if~exist(DIR_C,'dir'),mkdir(DIR_C);end
DIR_D=fullfile(OUT_ROOT,'情形D');if~exist(DIR_D,'dir'),mkdir(DIR_D);end
DIR_ALL=fullfile(OUT_ROOT,'总体（严谨三段）');if~exist(DIR_ALL,'dir'),mkdir(DIR_ALL);end
XLSX_MAIN=fullfile(pwd,'q3_result.xlsx');
TOPK_CASE=16;TOPK_ALL=16;SAVE_DPI=240;VERBOSE=true;
cnFont=chooseChineseFont();
set(groot,'defaultAxesFontName',cnFont,'defaultTextFontName',cnFont,...
    'defaultAxesFontSize',11,'defaultTextFontSize',10,'defaultFigureColor','w');
try
    set(groot,'defaultAxesToolbarVisible','off');
catch
end
%参数（A/B/C 都参与半段循环）
P=struct();
P.p_part=1-0.10*[1 1 1 1 1 1 1 1];
P.c_buy=[2 8 12 2 8 12 8 12];
P.c_test_part=[1 1 2 1 1 2 1 2];
P.y_half=[0.9 0.9 0.9];
P.c_asm_half=[8 8 8];
P.c_test_half=[4 4 4];
P.c_dis_half=[6 6 6];
P.y_final=0.9;P.c_asm_final=8;P.c_test_final=6;P.c_dis_final=10;
P.price=200;P.exch_loss=40;
P.aftersale_rework_with_test=true;
%分情形 A/B/C（各自Top16）
[T_A,compA,bitsA]=enumerate_half_case(P,[1 2 3],'A');
[T_B,compB,bitsB]=enumerate_half_case(P,[4 5 6],'B');
[T_C,compC,bitsC]=enumerate_half_case(P,[7 8],'C');
dump_case_results('情形A',T_A,compA,bitsA,XLSX_MAIN,DIR_AB,TOPK_CASE,SAVE_DPI,VERBOSE);
dump_case_results('情形B',T_B,compB,bitsB,XLSX_MAIN,DIR_AB,TOPK_CASE,SAVE_DPI,VERBOSE);
dump_case_results('情形C',T_C,compC,bitsC,XLSX_MAIN,DIR_C,TOPK_CASE,SAVE_DPI,VERBOSE);
%情形 D（成品层）
[T_D,compD,bitsD]=enumerate_final_case(P);
dump_case_results('情形D',T_D,compD,bitsD,XLSX_MAIN,DIR_D,TOPK_CASE,SAVE_DPI,VERBOSE);
%总体4096（利润降序）
[T_all,compAll,bitsAll]=search_all_policies_g(P);
TsortAll=sortrows(T_all,'profit','descend');
Kall=min(TOPK_ALL,height(TsortAll));
idxTopAll=TsortAll.index(1:Kall);
hdr_cn_all={'策略','单位成本','单位利润','一次出货合格率q',...
    'A段期望尝试N','B段期望尝试N','C段期望尝试N','成层期望尝试N',...
    '首发制造成本','返修期望成本','退换赔付期望','成层拆解费期望','位图(12位)'};
C_all=[TsortAll.policy,num2cell(TsortAll.cost),num2cell(TsortAll.profit),num2cell(TsortAll.q_ship),...
    num2cell(TsortAll.N_half1),num2cell(TsortAll.N_half2),num2cell(TsortAll.N_half3),num2cell(TsortAll.N_final),...
    num2cell(compAll.base(TsortAll.index)),num2cell(compAll.rework(TsortAll.index)),...
    num2cell(compAll.return_loss(TsortAll.index)),num2cell(compAll.final_dis(TsortAll.index)),...
    cellfun(@(v)sprintf('%d',v),num2cell(bitsAll(TsortAll.index,:),2),'uni',0)];
write_xlsx_cn([hdr_cn_all;C_all],XLSX_MAIN,'总体4096（利润降序_严谨三段）');
if VERBOSE
    fprintf('\n====== 总体Top%d（按单位客户利润降序；三段口径） ======\n',Kall);
    fmt='%-52s | 成本=%7.3f | 利润=%7.3f | q=%.4f | N_A=%.4f | N_B=%.4f | N_C=%.4f | N_final=%.4f\n';
    for r=1:Kall
        i=TsortAll.index(r);
        fprintf(fmt,T_all.policy{i},T_all.cost(i),T_all.profit(i),T_all.q_ship(i),...
            T_all.N_half1(i),T_all.N_half2(i),T_all.N_half3(i),T_all.N_final(i));
    end
end
plot_overall_figs(TsortAll,compAll,bitsAll,Kall,DIR_ALL,SAVE_DPI);
bestBits=bitsAll(TsortAll.index(1),:);
sensTbl=tornado_sensitivity(P,bestBits,DIR_ALL,SAVE_DPI);
write_xlsx_cn(sensTbl,XLSX_MAIN,'总体Top1_龙卷风敏感性');
fprintf('\n 核算：策略 11111111 | 半检=是 | 半拆=是 | 成检=否 | 成拆=是 \n');
bits_check=[ones(1,8),1 1,0 1];
o=eval_policy_bits_g(bits_check,P);
fprintf('成本=%.3f | 利润=%.3f | q=%.3f | N_A=%.3f N_B=%.3f N_C=%.3f N_final=%.3f\n',...
    o.cost_per_customer,o.profit_per_customer,o.q_ship,...
    o.n_half1_attempts,o.n_half2_attempts,o.n_half3_attempts,o.n_final_attempts);
fprintf('\n已完成所有输出。\nExcel：%s\n图像目录：%s\n',XLSX_MAIN,OUT_ROOT);
%===================== 函数区 =====================
function [Tcase,comp,bitsMat]=enumerate_half_case(P,idxParts,tag)
K=numel(idxParts);
nPol=2^(K+2);
pol_str=cell(nPol,1);
cost_arr=zeros(nPol,1);
q_arr=zeros(nPol,1);
N_arr=zeros(nPol,1);
b_base=zeros(nPol,1);
b_rewk=zeros(nPol,1);
b_dis=zeros(nPol,1);
bitsMat=zeros(nPol,K+2);
for s=0:nPol-1
    bits=bitget(uint16(s),K+2:-1:1);
    bitsMat(s+1,:)=bits;
    insp_part=false(1,K);insp_part(:)=bits(1:K);
    half_insp=bits(K+1)==1;
    half_dis=bits(K+2)==1;
    [H,sH,N_half,base1,rewk1,dis1]=half_block_case(P,idxParts,insp_part,half_insp,half_dis);
    cost_arr(s+1)=H;q_arr(s+1)=sH;N_arr(s+1)=N_half;
    b_base(s+1)=base1;b_rewk(s+1)=rewk1;b_dis(s+1)=dis1;
    pol_str{s+1}=half_bits2str(insp_part,half_insp,half_dis,idxParts,tag);
end
Tcase=table((1:nPol)',pol_str,cost_arr,q_arr,N_arr,...
    'VariableNames',{'index','policy','cost','q','N_attempt'});
comp=struct('base',b_base,'rework',b_rewk,'disasm',b_dis,'return_loss',zeros(nPol,1));
end
function [Hcost,sH,N_half,base1,rewk1,dis1]=half_block_case(P,idx,insp_part,half_insp,half_dis)
K=numel(idx);
p=P.p_part(idx);cb=P.c_buy(idx);ct=P.c_test_part(idx);
yH=P.y_half;if numel(yH)>1,yH=yH(1);end
cA=P.c_asm_half;if numel(cA)>1,cA=cA(1);end
cT=P.c_test_half;if numel(cT)>1,cT=cT(1);end
cD=P.c_dis_half;if numel(cD)>1,cD=cD(1);end
eps=1e-12;
r=ones(1,K);r(~insp_part)=p(~insp_part);
sH=prod(r)*yH;sH=max(sH,eps);N_half=1/sH;
C_ins_parts_per_attempt=sum(((cb+ct)./max(p,eps)).*insp_part);
C_unins_per_attempt=sum(cb(~insp_part));
C_attempt_no_test=C_ins_parts_per_attempt+C_unins_per_attempt+cA;
C_attempt_with_test=C_attempt_no_test+cT;
if half_insp
    F=N_half-1;
    if half_dis
        C_ins_once=sum(((cb+ct)./max(p,eps)).*insp_part);
        C_unins_each=sum(cb(~insp_part));
        Hcost=C_ins_once+N_half*(C_unins_each+cA+cT)+F*cD;
        base1=C_ins_once+(C_unins_each+cA+cT);
        rewk1=(N_half-1)*(C_unins_each+cA+cT);
        dis1=F*cD;
    else
        Hcost=N_half*(C_ins_parts_per_attempt+C_unins_per_attempt+cA+cT);
        base1=C_ins_parts_per_attempt+C_unins_per_attempt+cA+cT;
        rewk1=(N_half-1)*(C_ins_parts_per_attempt+C_unins_per_attempt+cA+cT);
        dis1=0;
    end
else
    Hcost=C_attempt_no_test;base1=C_attempt_no_test;rewk1=0;dis1=0;
end
end
function [Tcase,comp,bitsMat]=enumerate_final_case(P)
pol_str=cell(4,1);cost_arr=zeros(4,1);q_arr=zeros(4,1);N_arr=zeros(4,1);
b_base=zeros(4,1);b_rewk=zeros(4,1);b_dis=zeros(4,1);b_ret=zeros(4,1);bitsMat=zeros(4,2);
yF=P.y_final;cAF=P.c_asm_final;cTF=P.c_test_final;cDF=P.c_dis_final;L=P.exch_loss;eps=1e-12;
for s=0:3
    bits=bitget(uint8(s),2:-1:1);
    final_insp=bits(1)==1;final_dis=bits(2)==1;bitsMat(s+1,:)=bits;
    if final_insp
        Nf=1/max(yF,eps);q=1;
        if final_dis
            base=(cAF+cTF);rewk=(Nf-1)*(cAF+cTF);dis=(Nf-1)*cDF;cost=base+rewk+dis;
        else
            cost=Nf*(cAF+cTF);base=cost;rewk=0;dis=0;
        end
        ret=0;
    else
        if P.aftersale_rework_with_test
            E=(final_dis)*(cDF+cAF+cTF)/max(yF,eps)+(~final_dis)*(cAF+cTF)/max(yF,eps);
            base=cAF;rewk=(1-yF)*E;dis=0;ret=(1-yF)*L;cost=base+rewk+ret;q=yF;Nf=1+(1-yF);
        else
            base=cAF;rewk=(1/yF-1)*cAF;dis=0;ret=(1/yF-1)*L;cost=base+rewk+ret;q=yF;Nf=1/yF;
        end
    end
    cost_arr(s+1)=cost;q_arr(s+1)=q;N_arr(s+1)=Nf;
    b_base(s+1)=base;b_rewk(s+1)=rewk;b_dis(s+1)=dis;b_ret(s+1)=ret;
    pol_str{s+1}=sprintf('成检=%s | 成拆=%s',yn(final_insp),yn(final_dis));
end
Tcase=table((1:4)',pol_str,cost_arr,q_arr,N_arr,...
    'VariableNames',{'index','policy','cost','q','N_attempt'});
comp=struct('base',b_base,'rework',b_rewk,'disasm',b_dis,'return_loss',b_ret);
bitsMat=bitsMat;
end
function dump_case_results(caseName,Tcase,comp,bitsMat,xlsx,outDir,TOPK,DPI,VERBOSE)
[~,ord]=sort(Tcase.cost,'ascend');Tsort=Tcase(ord,:);
compS.base=comp.base(ord);compS.rework=comp.rework(ord);
compS.disasm=comp.disasm(ord);compS.return_loss=comp.return_loss(ord);
bitsS=bitsMat(ord,:);
hdr_all={'策略','期望成本','一次成功率q','期望尝试次数N','首发成本','返修期望','拆解期望','赔付期望','位图'};
C_all=[Tsort.policy,num2cell(Tsort.cost),num2cell(Tsort.q),num2cell(Tsort.N_attempt),...
    num2cell(compS.base),num2cell(compS.rework),num2cell(compS.disasm),num2cell(compS.return_loss),...
    cellfun(@(v)sprintf('%d',v),num2cell(bitsS,2),'uni',0)];
write_xlsx_cn([hdr_all;C_all],xlsx,[caseName '（全部策略_按成本升序）']);
K=min(TOPK,height(Tsort));Ttop=Tsort(1:K,:);
compTop.base=compS.base(1:K);compTop.rework=compS.rework(1:K);
compTop.disasm=compS.disasm(1:K);compTop.return_loss=compS.return_loss(1:K);
bitsTop=bitsS(1:K,:);
C_top=[Ttop.policy,num2cell(Ttop.cost),num2cell(Ttop.q),num2cell(Ttop.N_attempt),...
    num2cell(compTop.base),num2cell(compTop.rework),num2cell(compTop.disasm),num2cell(compTop.return_loss),...
    cellfun(@(v)sprintf('%d',v),num2cell(bitsTop,2),'uni',0)];
write_xlsx_cn([hdr_all;C_top],xlsx,[caseName ' Top16（成本最低）']);
if VERBOSE
    fprintf('\n====%s：Top%d（成本最低）====\n',caseName,K);
    fmt='%-46s | 成本=%7.3f | q=%6.4f | N=%5.2f | 基=%7.3f | 返=%7.3f | 拆=%6.3f | 赔=%6.3f\n';
    for r=1:K
        fprintf(fmt,Ttop.policy{r},Ttop.cost(r),Ttop.q(r),Ttop.N_attempt(r),...
            compTop.base(r),compTop.rework(r),compTop.disasm(r),compTop.return_loss(r));
    end
end
f1=figure('Color','w','Position',[100 100 1180 520]);
bar(Ttop.cost);grid on;ylabel('期望成本');title([caseName '：Top16 成本（越低越好）']);
set(gca,'XTick',1:K,'XTickLabel',Ttop.policy,'XTickLabelRotation',25);
saveFig_noToolbar(f1,fullfile(outDir,[caseName '_Top16_成本条形.png']),DPI);
f2=figure('Color','w','Position',[100 100 1320 560]);
M=[compTop.base,compTop.rework,compTop.disasm,compTop.return_loss];
bar(categorical(1:K),M,'stacked');grid on;ylabel('期望成本');
title([caseName '：Top16 成本构成']);legend({'首发','返修','拆解','赔付'},'Location','best');
set(gca,'XTickLabel',Ttop.policy,'XTickLabelRotation',25);
saveFig_noToolbar(f2,fullfile(outDir,[caseName '_Top16_成本堆叠.png']),DPI);
f3=figure('Color','w','Position',[100 100 980 520]);
scatter(Tsort.q,Tsort.cost,22,'filled');grid on;
xlabel('一次成功率 q');ylabel('期望成本');title([caseName '：成本 vs q（全部策略）']);
saveFig_noToolbar(f3,fullfile(outDir,[caseName '_全体_成本_vs_q.png']),DPI);
f4=figure('Color','w','Position',[100 100 880 520]);
imagesc(bitsTop);set(gca,'YDir','normal');caxis([0 1]);
colormap([0.98 0.85 0.60;0.16 0.45 0.85]);colorbar('Ticks',[0 1],'TickLabels',{'关(0)','开(1)'});
title([caseName '：Top16 位图（1=开,0=关）']);xlabel('位索引');ylabel('Top序位');
set(gca,'XTick',1:size(bitsTop,2),'YTick',1:size(bitsTop,1));
[nr,nc]=size(bitsTop);
for r=1:nr
    for c=1:nc
        val=bitsTop(r,c);
        txtColor=val*[1 1 1]+(1-val)*[0 0 0];
        text(c,r,num2str(val),'Horiz','center','Vert','middle','FontSize',9,'FontWeight','bold','Color',txtColor);
    end
end
saveFig_noToolbar(f4,fullfile(outDir,[caseName '_Top16_位图热力.png']),DPI);
end
function [T,comp,bitsMat]=search_all_policies_g(P)
nPol=4096;
pol_bits=cell(nPol,1);cost_arr=zeros(nPol,1);profit_arr=zeros(nPol,1);qship_arr=zeros(nPol,1);
Nh1_arr=zeros(nPol,1);Nh2_arr=zeros(nPol,1);Nh3_arr=zeros(nPol,1);Nfin_arr=zeros(nPol,1);
base_arr=zeros(nPol,1);rewk_arr=zeros(nPol,1);ret_arr=zeros(nPol,1);fdis_arr=zeros(nPol,1);
bitsMat=zeros(nPol,12);
for k=0:nPol-1
    bits=bitget(uint16(k),12:-1:1);bitsMat(k+1,:)=bits;
    out=eval_policy_bits_g(bits,P);
    pol_bits{k+1}=bits2str_g(bits);
    cost_arr(k+1)=out.cost_per_customer;
    profit_arr(k+1)=out.profit_per_customer;
    qship_arr(k+1)=out.q_ship;
    Nh1_arr(k+1)=out.n_half1_attempts;
    Nh2_arr(k+1)=out.n_half2_attempts;
    Nh3_arr(k+1)=out.n_half3_attempts;
    Nfin_arr(k+1)=out.n_final_attempts;
    base_arr(k+1)=out.comp_base;rewk_arr(k+1)=out.comp_rework;
    ret_arr(k+1)=out.comp_return_loss;fdis_arr(k+1)=out.comp_final_dis;
end
T=table((1:nPol)',pol_bits,cost_arr,profit_arr,qship_arr,Nh1_arr,Nh2_arr,Nh3_arr,Nfin_arr,...
    'VariableNames',{'index','policy','cost','profit','q_ship','N_half1','N_half2','N_half3','N_final'});
comp=struct('base',base_arr,'rework',rewk_arr,'return_loss',ret_arr,'final_dis',fdis_arr);
end
function out=eval_policy_bits_g(bits,P)
idxA=[1 2 3];idxB=[4 5 6];idxC=[7 8];
insp_part=logical(bits(1:8));half_insp=bits(9)==1;half_dis=bits(10)==1;
final_insp=bits(11)==1;final_dis=bits(12)==1;
p=P.p_part(:)';cb=P.c_buy(:)';ct=P.c_test_part(:)';
yH=P.y_half(:)';cAH=P.c_asm_half(:)';cTH=P.c_test_half(:)';cDH=P.c_dis_half(:)';
yF=P.y_final;cAF=P.c_asm_final;cTF=P.c_test_final;cDF=P.c_dis_final;
PRICE=P.price;LOSS=P.exch_loss;eps=1e-12;
[H1,sH1,N1]=half_block_g(idxA,1,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
[H2,sH2,N2]=half_block_g(idxB,2,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
[H3,sH3,N3]=half_block_g(idxC,3,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
comp_base=0;comp_rework=0;comp_return=0;comp_fdis=0;
if final_insp
    if half_insp
        Nf=1/max(yF,eps);
        if final_dis
            comp_base=(H1.C_good+H2.C_good+H3.C_good)+(cAF+cTF);
            comp_rework=(Nf-1)*(cAF+cTF);
            comp_fdis=(Nf-1)*cDF;
            cost=comp_base+comp_rework+comp_fdis;
        else
            cost=Nf*((H1.C_good+H2.C_good+H3.C_good)+cAF+cTF);
            comp_base=cost;comp_rework=0;comp_fdis=0;
        end
        q_ship=1;N_final=Nf;
    else
        s_final=sH1*sH2*sH3*yF;Nf=1/max(s_final,eps);
        cost_once=(H1.C_attempt+H2.C_attempt+H3.C_attempt+cAF+cTF);
        if final_dis
            comp_base=cost_once;
            comp_rework=(Nf-1)*cost_once;
            comp_fdis=(Nf-1)*cDF;
        else
            comp_base=Nf*cost_once;comp_rework=0;comp_fdis=0;
        end
        cost=comp_base+comp_rework+comp_fdis;q_ship=1;N_final=Nf;
    end
    comp_return=0;
    cost_per_customer=cost;
    profit_per_customer=PRICE-cost_per_customer;
else
    if half_insp
        base_cost=(H1.C_good+H2.C_good+H3.C_good+cAF);
        if isfield(P,'aftersale_rework_with_test')&&P.aftersale_rework_with_test
            if final_dis
                E_rework=(cDF+cAF+cTF)/max(yF,eps);
            else
                E_rework=((H1.C_good+H2.C_good+H3.C_good)+cAF+cTF)/max(yF,eps);
            end
            comp_base=base_cost;
            comp_rework=(1-yF)*E_rework;
            comp_return=(1-yF)*LOSS;
            comp_fdis=0;
            cost_per_customer=comp_base+comp_rework+comp_return;
            profit_per_customer=PRICE-cost_per_customer;
            q_ship=yF;N_final=1+(1-yF);
        else
            s_ship=max(yF,eps);N_ship=1/s_ship;
            cost_once=(H1.C_good+H2.C_good+H3.C_good+cAF);
            comp_base=cost_once;comp_rework=(N_ship-1)*cost_once;comp_return=(N_ship-1)*LOSS;
            comp_fdis=0;cost_per_customer=comp_base+comp_rework+comp_return;
            profit_per_customer=PRICE-cost_per_customer;q_ship=s_ship;N_final=N_ship;
        end
    else
        s_ship=max(sH1*sH2*sH3*yF,eps);N_ship=1/s_ship;
        cost_once=(H1.C_attempt+H2.C_attempt+H3.C_attempt+cAF);
        comp_base=cost_once;comp_rework=(N_ship-1)*cost_once;comp_return=(N_ship-1)*LOSS;
        comp_fdis=0;cost_per_customer=comp_base+comp_rework+comp_return;
        profit_per_customer=PRICE-cost_per_customer;q_ship=s_ship;N_final=N_ship;
    end
end
out.cost_per_customer=cost_per_customer;
out.profit_per_customer=profit_per_customer;
out.q_ship=q_ship;
out.n_half1_attempts=N1;
out.n_half2_attempts=N2;
out.n_half3_attempts=N3;
out.n_final_attempts=N_final;
out.comp_base=comp_base;
out.comp_rework=comp_rework;
out.comp_return_loss=comp_return;
out.comp_final_dis=comp_fdis;
end
function [H,sH,N_half]=half_block_g(idx,hID,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps)
if numel(yH)<3,yH=repmat(yH(1),1,3);end
if numel(cAH)<3,cAH=repmat(cAH(1),1,3);end
if numel(cTH)<3,cTH=repmat(cTH(1),1,3);end
if numel(cDH)<3,cDH=repmat(cDH(1),1,3);end
r=ones(1,numel(idx));r(~insp_part(idx))=p(idx(~insp_part(idx)));
sH=prod(r)*yH(hID);sH=max(sH,eps);N_half=1/sH;
C_ins_parts_per_attempt=sum(((cb(idx)+ct(idx))./max(p(idx),eps)).*insp_part(idx));
C_unins_per_attempt=sum(cb(idx(~insp_part(idx))));
C_attempt_no_test=C_ins_parts_per_attempt+C_unins_per_attempt+cAH(hID);
C_attempt_with_test=C_attempt_no_test+cTH(hID);
H=struct();
if half_insp
    F=N_half-1;
    if half_dis
        C_ins_once=sum(((cb(idx)+ct(idx))./max(p(idx),eps)).*insp_part(idx));
        C_unins_each=sum(cb(idx(~insp_part(idx))));
        H.C_good=C_ins_once+N_half*(C_unins_each+cAH(hID)+cTH(hID))+F*cDH(hID);
    else
        H.C_good=N_half*(C_ins_parts_per_attempt+C_unins_per_attempt+cAH(hID)+cTH(hID));
    end
    H.C_attempt=C_attempt_with_test;
else
    H.C_good=C_attempt_no_test;
    H.C_attempt=C_attempt_no_test;
    N_half=1;
end
end
function plot_overall_figs(TsortAll,compAll,bitsAll,K,outDir,DPI)
if~exist(outDir,'dir'),mkdir(outDir);end
f1=figure('Color','w','Position',[80 80 1200 560]);
scatter(TsortAll.cost,TsortAll.profit,18,'filled');grid on;
xlabel('单位成本');ylabel('单位利润');title('总体：成本–利润散点（全部4096）');
hold on;scatter(TsortAll.cost(1:K),TsortAll.profit(1:K),36);
text(TsortAll.cost(1:K),TsortAll.profit(1:K),strcat("  ",string(1:K)),'FontSize',9);
hold off;saveFig_noToolbar(f1,fullfile(outDir,'总体_成本vs利润_散点.png'),DPI);
f2=figure('Color','w','Position',[80 80 1320 580]);
idxTop=TsortAll.index(1:K);
M=[compAll.base(idxTop),compAll.rework(idxTop),compAll.return_loss(idxTop),compAll.final_dis(idxTop)];
bar(categorical(1:K),M,'stacked');grid on;ylabel('单位成本');
title('总体Top16：成本构成');legend({'首发','返修','赔付','成拆'},'Location','best');
set(gca,'XTickLabel',TsortAll.policy(1:K),'XTickLabelRotation',25);
saveFig_noToolbar(f2,fullfile(outDir,'总体Top16_成本堆叠.png'),DPI);
f3=figure('Color','w','Position',[80 80 920 540]);
bitsTop=bitsAll(idxTop,:);
imagesc(bitsTop);set(gca,'YDir','normal');caxis([0 1]);
colormap([0.98 0.85 0.60;0.16 0.45 0.85]);colorbar('Ticks',[0 1],'TickLabels',{'关(0)','开(1)'});
title('总体Top16：位图（1=开,0=关）');xlabel('位索引');ylabel('Top序位');
set(gca,'XTick',1:size(bitsTop,2),'YTick',1:size(bitsTop,1));
[nr,nc]=size(bitsTop);
for r=1:nr
    for c=1:nc
        v=bitsTop(r,c);
        txtColor=v*[1 1 1]+(1-v)*[0 0 0];
        text(c,r,num2str(v),'Horiz','center','Vert','middle','FontSize',9,'FontWeight','bold','Color',txtColor);
    end
end
saveFig_noToolbar(f3,fullfile(outDir,'总体Top16_位图热力.png'),DPI);
f4=figure('Color','w','Position',[80 80 980 480]);
freq=mean(bitsTop,1)*100;
bar(freq);grid on;xlabel('位索引');ylabel('开启频率(%)');title('总体Top16：位开启频率');
saveFig_noToolbar(f4,fullfile(outDir,'总体Top16_位开启频率.png'),DPI);
f5=figure('Color','w','Position',[80 80 980 480]);
histogram(TsortAll.profit,40);grid on;xlabel('单位利润');ylabel('策略数');
title('总体：单位利润分布（4096策略）');
saveFig_noToolbar(f5,fullfile(outDir,'总体_利润直方图.png'),DPI);
end
function sensTbl=tornado_sensitivity(P,bits,outDir,DPI)
base=eval_policy_bits_g(bits,P);
baseProfit=base.profit_per_customer;
spec={...
    'S_total(售价)','price',0.10,'rel';...
    '退换损失','exch_loss',0.10,'rel';...
    '成装','c_asm_final',0.10,'rel';...
    '成检','c_test_final',0.10,'rel';...
    '成拆','c_dis_final',0.10,'rel';...
    '半装(A/B/C)','c_asm_half',0.10,'rel';...
    '半检(A/B/C)','c_test_half',0.10,'rel';...
    '半拆(A/B/C)','c_dis_half',0.10,'rel';...
    '零件采购(均匀+10%)','c_buy',0.10,'rel';...
    '来料检费(均匀+10%)','c_test_part',0.10,'rel';...
    '零件合格率p(±0.02)','p_part',0.02,'abs';...
    '半段良率yH(±0.02)','y_half',0.02,'abs';...
    '成段良率yF(±0.02)','y_final',0.02,'abs'...
    };
n=size(spec,1);low=zeros(n,1);high=zeros(n,1);
labels=spec(:,1);
for i=1:n
    field=spec{i,2};delta=spec{i,3};mode=spec{i,4};
    P1=P;P2=P;
    switch field
        case {'price','exch_loss','c_asm_final','c_test_final','c_dis_final'}
            if strcmp(mode,'rel'),P1.(field)=P.(field)*(1-delta);P2.(field)=P.(field)*(1+delta);
            else,P1.(field)=P.(field)-delta;P2.(field)=P.(field)+delta;end
        case {'c_asm_half','c_test_half','c_dis_half','c_buy','c_test_part'}
            if strcmp(mode,'rel'),P1.(field)=P.(field)*(1-delta);P2.(field)=P.(field)*(1+delta);
            else,P1.(field)=P.(field)-delta;P2.(field)=P.(field)+delta;end
        case {'p_part','y_half'}
            P1.(field)=max(0,min(1,P.(field)-delta));P2.(field)=max(0,min(1,P.(field)+delta));
        case 'y_final'
            P1.(field)=max(0,min(1,P.(field)-delta));P2.(field)=max(0,min(1,P.(field)+delta));
    end
    low(i)=eval_policy_bits_g(bits,P1).profit_per_customer-baseProfit;
    high(i)=eval_policy_bits_g(bits,P2).profit_per_customer-baseProfit;
end
sensTbl=table((1:n)',string(labels),low,high,'VariableNames',{'Rank','Factor','DeltaMinus','DeltaPlus'});
[~,ord]=sort(max(abs([low,high]),[],2),'descend');sensTbl=sensTbl(ord,:);
f=figure('Color','w','Position',[120 120 900 640]);
Y=categorical(sensTbl.Factor);Y=reordercats(Y,sensTbl.Factor);
M=[sensTbl.DeltaMinus,sensTbl.DeltaPlus];
barh(Y,M,'stacked');grid on;
xlabel('Δ利润（相对基线，元/件）');title('Top策略龙卷风敏感性（±10%或 ±0.02）');
legend({'-变动','+变动'},'Location','best');
saveFig_noToolbar(f,fullfile(outDir,'总体Top1_龙卷风敏感性.png'),DPI);
end
function s=half_bits2str(insp_part,half_insp,half_dis,idxParts,tag)
bitsC=sprintf('%d',insp_part);
s=sprintf('(%s)零[%s] | 半检=%s | 半拆=%s',tag,bitsC,yn(half_insp),yn(half_dis));
end
function s=bits2str_g(bits)
s=sprintf('%s | 半检=%s | 半拆=%s | 成检=%s | 成拆=%s',...
    sprintf('%d',bits(1:8)),yn(bits(9)),yn(bits(10)),yn(bits(11)),yn(bits(12)));
end
function y=yn(b)
if b~=0,y='是';else,y='否';end
end
function write_xlsx_cn(cell_or_table,xlsx_file,sheetname)
sheet=regexprep(sheetname,'[:\\/\?\*\[\]]','_');
if istable(cell_or_table)
    try
        writetable(cell_or_table,xlsx_file,'Sheet',sheet,'WriteMode','overwritesheet');
    catch
        writetable(cell_or_table,xlsx_file,'Sheet',sheet);
    end
else
    try
        writecell(cell_or_table,xlsx_file,'Sheet',sheet);
    catch
        C=cell_or_table;C(cellfun(@isempty,C))={''};
        try
            xlswrite(xlsx_file,C,sheet);
        catch
            Ttmp=cell2table(C(2:end,:));
            writetable(Ttmp,xlsx_file,'Sheet',sheet);
        end
    end
end
end

function saveFig_noToolbar(f,filename,dpi)
ax=findall(f,'type','axes');
for ii=1:numel(ax)
    try
        ax(ii).Toolbar.Visible='off';
    catch
        try
            axtoolbar(ax(ii),'visible','off');
        catch
        end
    end
end
drawnow;
try
    exportgraphics(f,filename,'Resolution',dpi);
catch
    [p,~,~]=fileparts(filename);if~exist(p,'dir'),mkdir(p);end
    print(f,filename,'-dpng',sprintf('-r%d',dpi));
end
end

function fname=chooseChineseFont()
cands={'Microsoft YaHei','SimHei','PingFang SC','Heiti SC','Songti SC','Noto Sans CJK SC','Source Han Sans SC','Hiragino Sans GB'};
avail=listfonts;fname=get(groot,'defaultAxesFontName');
for i=1:numel(cands)
    if any(strcmp(avail,cands{i})),fname=cands{i};return;end
end
end
