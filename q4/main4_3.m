%main4_3.m 问题4：稳健置信集（悲观点）对照，可调用 main4_1 中函数
clc;clear;close all;
%基本输出 
OUT_ROOT=fullfile(pwd,'q4_output');
XLSX=fullfile(OUT_ROOT,'q4_results.xlsx');
DIR_Q2=fullfile(OUT_ROOT,'Q2_risk');
DIR_Q3=fullfile(OUT_ROOT,'Q3_risk');
SAVE_DPI=240;TOPK=16;
%先验与分位
prior=[0.5,0.5];%Jeffreys
qTail=0.95;%缺陷率单侧95%→ 良率悲观= 1-betainv(0.95,...)
%优先使用主工程函数；如果没有就启用本地备用辅函数 
evalPolicy_F=pickFunc(@eval_policy_bits_g,@eval_policy_bits_g_local);
bits2str_F=pickFunc(@bits2str_g,@bits2str_g_local);
saveFig_F=pickFunc(@saveFig_noToolbar,@saveFig_noToolbar_local);
font_F=pickFunc(@chooseChineseFont,@chooseChineseFont_local);
writeTbl_F=pickFunc(@writetable_safe,@writetable_safe_local);
loadS_F=pickFunc(@load_sampling_or_default,@load_sampling_or_default_local);
%外观 
cnFont=font_F();
set(groot,'defaultAxesFontName',cnFont,'defaultTextFontName',cnFont,...
          'defaultAxesFontSize',11,'defaultTextFontSize',10,'defaultFigureColor','w');
try set(groot,'defaultAxesToolbarVisible','off');catch,end
%读取抽样/参数（与 main4_1 一致） 
S=loadS_F('q4_input.xlsx');
%Q2：16策略悲观点（I1/I2/IF/D）
pol16=dec2bin(0:15,4)-'0';
names16=arrayfun(@(i) strategyNameQ2(pol16(i,:)),1:16,'uni',0);
p1_lo=1-betainv(qTail,prior(1)+S.Q2_parts.x(1),prior(2)+S.Q2_parts.n(1)-S.Q2_parts.x(1));
p2_lo=1-betainv(qTail,prior(1)+S.Q2_parts.x(2),prior(2)+S.Q2_parts.n(2)-S.Q2_parts.x(2));
yF_lo=1-betainv(qTail,prior(1)+S.Q2_final.x(1),prior(2)+S.Q2_final.n(1)-S.Q2_final.x(1));
cb1=S.Q2_parts.c_buy(1);ct1=S.Q2_parts.c_test(1);
cb2=S.Q2_parts.c_buy(2);ct2=S.Q2_parts.c_test(2);
cAF=S.Q2_final.c_AF(1);cTF=S.Q2_final.c_TF(1);cDF=S.Q2_final.c_DF(1);
L=S.Q2_final.L(1);S_total=S.Q2_final.S_total(1);
prof_rob16=zeros(16,1);
for k=1:16
    bits=pol16(k,:);
    prof_rob16(k)=q2_profit_det(bits,p1_lo,p2_lo,yF_lo,[cb1,ct1,cb2,ct2],cAF,cTF,cDF,L,S_total);
end
[~,ord16]=sort(prof_rob16,'descend');
Tq2=table(names16(:),prof_rob16,'VariableNames',{'策略','悲观点利润'});
writeTbl_F(Tq2,XLSX,'Q2_稳健对照');
fQ2=figure('Color','w','Position',[80 80 1000 560]);
bar(prof_rob16(ord16));grid on;ylabel('悲观点利润(元/单)');
title('Q4–Q2：悲观点（单侧95%）Top16');
set(gca,'XTick',1:16,'XTickLabel',names16(ord16),'XTickLabelRotation',20);
saveFig_F(fQ2,fullfile(DIR_Q2,'Q2_robust_pessimistic.png'),SAVE_DPI);
%Q3：4096策略悲观点
nPol=4096;
bitsMat= logical(dec2bin(0:nPol-1,12)-'0');
%常量参数（悲观良率）
Pconst=struct();
Pconst.p_part=1-betainv(qTail,prior(1)+S.Q3_parts.x',prior(2)+S.Q3_parts.n'-S.Q3_parts.x')';
Pconst.y_half=1-betainv(qTail,prior(1)+S.Q3_half.x',prior(2)+S.Q3_half.n' -S.Q3_half.x')';
Pconst.y_final=1-betainv(qTail,prior(1)+S.Q3_final.x(1),prior(2)+S.Q3_final.n(1)- S.Q3_final.x(1));
Pconst.c_buy=S.Q3_parts.c_buy';
Pconst.c_test_part=S.Q3_parts.c_test';
Pconst.c_asm_half=S.Q3_half.cA';Pconst.c_test_half=S.Q3_half.cT';Pconst.c_dis_half=S.Q3_half.cD';
Pconst.c_asm_final=S.Q3_final.c_AF(1);Pconst.c_test_final=S.Q3_final.c_TF(1);Pconst.c_dis_final=S.Q3_final.c_DF(1);
Pconst.price=S.Q3_final.S_total(1);
Pconst.exch_loss=S.Q3_final.L(1);
Pconst.aftersale_rework_with_test=true;
prof_rob3=zeros(nPol,1);
parfor i=1:nPol
    b=bitsMat(i,:);
    out=evalPolicy_F(b,Pconst);
    prof_rob3(i)=out.profit_per_customer;
end
[~,ord3]=sort(prof_rob3,'descend');keep=min(TOPK,nPol);
pol_str=arrayfun(@(i) bits2str_F(bitsMat(i,:)),1:nPol,'uni',0)';
Tq3=table((1:nPol)',pol_str,prof_rob3,'VariableNames',{'index','policy','悲观点利润'});
writeTbl_F(Tq3,XLSX,'Q3_稳健对照');
fQ3=figure('Color','w','Position',[80 80 1200 560]);
bar(prof_rob3(ord3(1:keep)));grid on;ylabel('悲观点利润(元/单)');
title(sprintf('Q4–Q3：悲观点（单侧95%%）Top%d',keep));
set(gca,'XTick',1:keep,'XTickLabel',pol_str(ord3(1:keep)),'XTickLabelRotation',20);
saveFig_F(fQ3,fullfile(DIR_Q3,'Q3_robust_pessimistic_TopK.png'),SAVE_DPI);
fprintf('\n main4_3 完成。输出目录：%s\nExcel：%s\n',OUT_ROOT,XLSX);
%辅助函数区
function F=pickFunc(pref,fallback)
%若首选函数在路径上存在，则返回其句柄；否则返回备用
    if exist(func2str(pref),'file')
        F=pref;
    else
        F=fallback;
    end
end
%Q2：确定性悲观点利润 
function val=q2_profit_det(bits,p1,p2,yF,costP,cAF,cTF,cDF,L,S_total)
    I1=bits(1);I2=bits(2);IF=bits(3);D=bits(4);
    cb1=costP(1);ct1=costP(2);cb2=costP(3);ct2=costP(4);
    C_ins=I1*((cb1+ct1)/max(p1,realmin))+I2*((cb2+ct2)/max(p2,realmin));
    C_un=(1-I1)*cb1+(1-I2)*cb2;
    q_ship=(I1+(1-I1)*p1) * (I2+(1-I2)*p2) * yF;
    if IF==1
        EN=1/max(q_ship,realmin);
        cost=C_ins+EN*(C_un+cAF+cTF)+(EN-1)*D*cDF;
        val=S_total-cost;
    else
        base=C_ins+C_un+cAF;
        ratio=(1-q_ship)/max(q_ship,realmin);
        crework=ratio * (cDF+cAF+cTF);
        cret=ratio * L;
        cost=base+crework+cret;
        val=S_total-cost;
    end
end
function s=strategyNameQ2(bits)
    yn=@(b) ternary(b,'是','否');
    s=sprintf('预检1=%s，预检2=%s，终检=%s，拆=%s',yn(bits(1)),yn(bits(2)),yn(bits(3)),yn(bits(4)));
end
function y=ternary(c,a,b),if c,y=a;else,y=b;end,end
%备用辅助函数 
function out=eval_policy_bits_g_local(bits,P)
    idxA=[1 2 3];idxB=[4 5 6];idxC=[7 8];
    insp_part=logical(bits(1:8));
    half_insp=bits(9)==1;half_dis=bits(10)==1;
    final_insp= bits(11)==1;final_dis=bits(12)==1;
    p=P.p_part(:)';cb=P.c_buy(:)';ct=P.c_test_part(:)';
    yH=P.y_half(:)';cAH=P.c_asm_half(:)';cTH=P.c_test_half(:)';cDH=P.c_dis_half(:)';
    yF=P.y_final;cAF=P.c_asm_final;cTF=P.c_test_final;cDF=P.c_dis_final;
    PRICE=P.price;LOSS=P.exch_loss;eps=1e-12;
    [H1,sH1,N1]=half_block_g_local(idxA,1,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
    [H2,sH2,N2]=half_block_g_local(idxB,2,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
    [H3,sH3,N3]=half_block_g_local(idxC,3,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps);
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
                comp_base=cost_once;comp_rework=(Nf-1)*cost_once;comp_fdis=(Nf-1)*cDF;
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
            if isfield(P,'aftersale_rework_with_test') && P.aftersale_rework_with_test
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
                comp_fdis=0;
                cost_per_customer=comp_base+comp_rework+comp_return;
                profit_per_customer=PRICE-cost_per_customer;
                q_ship=s_ship;N_final=N_ship;
            end
        else
            s_ship=max(sH1*sH2*sH3*yF,eps);N_ship=1/s_ship;
            cost_once=(H1.C_attempt+H2.C_attempt+H3.C_attempt+cAF);
            comp_base=cost_once;comp_rework=(N_ship-1)*cost_once;comp_return=(N_ship-1)*LOSS;
            comp_fdis=0;
            cost_per_customer=comp_base+comp_rework+comp_return;
            profit_per_customer=PRICE-cost_per_customer;q_ship=s_ship;N_final=N_ship;
        end
    end
    out.cost_per_customer=cost_per_customer;out.profit_per_customer=profit_per_customer;
    out.q_ship=q_ship;out.n_half1_attempts=N1;out.n_half2_attempts=N2;out.n_half3_attempts=N3;out.n_final_attempts=N_final;
    out.comp_base=comp_base;out.comp_rework=comp_rework;out.comp_return_loss=comp_return;out.comp_final_dis=comp_fdis;
end
function [H,sH,N_half]=half_block_g_local(idx,hID,insp_part,half_insp,half_dis,p,cb,ct,yH,cAH,cTH,cDH,eps)
    if numel(yH)<3,yH=repmat(yH(1),1,3);end
    if numel(cAH)<3,cAH=repmat(cAH(1),1,3);end
    if numel(cTH)<3,cTH=repmat(cTH(1),1,3);end
    if numel(cDH)<3,cDH=repmat(cDH(1),1,3);end
    r=ones(1,numel(idx));bad=~insp_part(idx);r(bad)=p(idx(bad));
    sH=prod(r)*yH(hID);sH=max(sH,eps);N_half=1/sH;
    C_ins_parts=sum(((cb(idx)+ct(idx))./max(p(idx),eps)).*insp_part(idx));
    C_unins=sum(cb(idx(bad)));
    C_no_test=C_ins_parts+C_unins+cAH(hID);
    C_with_test=C_no_test+cTH(hID);
    H=struct();
    if half_insp
        F=N_half-1;
        if half_dis
            C_ins_once=sum(((cb(idx)+ct(idx))./max(p(idx),eps)).*insp_part(idx));
            C_unins_each=sum(cb(idx(bad)));
            H.C_good=C_ins_once+N_half*(C_unins_each+cAH(hID)+cTH(hID))+F*cDH(hID);
        else
            H.C_good=N_half*(C_ins_parts+C_unins+cAH(hID)+cTH(hID));
        end
        H.C_attempt=C_with_test;
    else
        H.C_good=C_no_test;H.C_attempt=C_no_test;N_half=1;
    end
end
function s=bits2str_g_local(bits)
    s=sprintf('%s | 半检=%s | 半拆=%s | 成检=%s | 成拆=%s',...
        sprintf('%d',bits(1:8)),yn(bits(9)),yn(bits(10)),yn(bits(11)),yn(bits(12)));
end
function y=yn(b),if b~=0,y='是';else,y='否';end,end
function fname=chooseChineseFont_local()
    cands={'Microsoft YaHei','SimHei','PingFang SC','Heiti SC','Songti SC','Noto Sans CJK SC','Source Han Sans SC','Hiragino Sans GB'};
    avail=listfonts;fname=get(groot,'defaultAxesFontName');
    for i=1:numel(cands)
        if any(strcmp(avail,cands{i})),fname=cands{i};return;end
    end
end
function saveFig_noToolbar_local(f,filename,dpi)
    ax=findall(f,'type','axes');
    for ii=1:numel(ax)
        try ax(ii).Toolbar.Visible='off';catch
            try axtoolbar(ax(ii),'visible','off');catch,end
        end
    end
    drawnow;
    [p,~,~]=fileparts(filename);
    if ~exist(p,'dir'),mkdir(p);end
    try
        exportgraphics(f,filename,'Resolution',dpi);
    catch
        print(f,filename,'-dpng',sprintf('-r%d',dpi));
    end
end
function writetable_safe_local(T,xlsx,sheetname)
    sheet=regexprep(sheetname,'[:\\/\?\*\[\]]','_');
    try
        writetable(T,xlsx,'Sheet',sheet,'WriteMode','overwritesheet');
    catch
        writetable(T,xlsx,'Sheet',sheet);
    end
end
function S=load_sampling_or_default_local(xlsx)
    S=struct();
    if exist(xlsx,'file')
        try
            S.Q2_parts=readtable(xlsx,'Sheet','Q2_parts');
            S.Q2_final=readtable(xlsx,'Sheet','Q2_final');
            S.Q3_parts=readtable(xlsx,'Sheet','Q3_parts');
            S.Q3_half=readtable(xlsx,'Sheet','Q3_half');
            S.Q3_final=readtable(xlsx,'Sheet','Q3_final');
            return;
        catch,warning('读取%s 失败，改用示例基线。',xlsx);
        end
    end
    S.Q2_parts=cell2table({'P1',100,10,6,2;'P2',100,10,10,2},...
        'VariableNames',{'name','n','x','c_buy','c_test'});
    S.Q2_final=table(100,10,8,6,10,10,56,...
        'VariableNames',{'n','x','c_AF','c_TF','c_DF','L','S_total'});
    names=compose('Part%d',1:8).';
    S.Q3_parts=table(names,repmat(100,8,1),repmat(10,8,1),...
        'VariableNames',{'name','n','x'});
    S.Q3_parts.c_buy=[2 8 12 2 8 12 8 12]';
    S.Q3_parts.c_test=[1 1  2 1 1  2 1  2]';
    S.Q3_half=cell2table({'A',100,10,8,4,6;'B',100,10,8,4,6;'C',100,10,8,4,6},...
        'VariableNames',{'name','n','x','cA','cT','cD'});
    S.Q3_final=table(100,10,8,6,10,40,200,...
        'VariableNames',{'n','x','c_AF','c_TF','c_DF','L','S_total'});
end