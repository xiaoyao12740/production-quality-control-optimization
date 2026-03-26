%main4_1.m-问题4：基于Q1抽样的后验Monte Carlo，重做Q2与Q3
clc;clear;close all;
tic;
%运行开关（若在 main4_2/4_3 先行设置，这里会沿用） 
if ~exist('RUN_Q2','var');RUN_Q2=true;end
if ~exist('RUN_Q3','var');RUN_Q3=true;end
%输出 
OUT_ROOT=fullfile(pwd,'q4_output');
if ~exist(OUT_ROOT,'dir');mkdir(OUT_ROOT);end
XLSX=fullfile(OUT_ROOT,'q4_results.xlsx');
SAVE_DPI=240;TOPK=16;
%字体与外观 
cnFont=chooseChineseFont();
set(groot,'defaultAxesFontName',cnFont,'defaultTextFontName',cnFont,...
          'defaultAxesFontSize',11,'defaultTextFontSize',10,'defaultFigureColor','w');
try
    set(groot,'defaultAxesToolbarVisible','off');
catch
end
%Monte Carlo 配置&先验 
M=2000;%后验样本条数
prior=[0.5,0.5];%Beta(1/2,1/2) Jeffreys先验
%读取抽样/参数（无文件时自动兜底示例） 
S=load_sampling_or_default('q4_input.xlsx');
%Q2：两件来料+成段（四位策略：I1/I2/IF/D） 
if RUN_Q2
    Q2=run_Q2_mc(S,prior,M,OUT_ROOT,XLSX,SAVE_DPI,TOPK);
end
%Q3：三段半品+成段（十二位策略，4096枚举） 
if RUN_Q3
    Q3=run_Q3_mc(S,prior,M,OUT_ROOT,XLSX,SAVE_DPI,TOPK);
end
fprintf('\n Q4 已完成：输出目录%s\n汇总Excel：%s\n',OUT_ROOT,XLSX);
t=toc;
fprintf('main4_1 实测耗时：%.2f 秒\n',t);
%函数区 
function S=load_sampling_or_default(xlsx)
%预期工作表Sheet：
%Q2_parts: name,n,x,c_buy,c_test
%Q2_final: n,x,c_AF,c_TF,c_DF,L,S_total
%Q3_parts: name(1..8),n,x,c_buy,c_test
%Q3_half : name(A,B,C),n,x,cA,cT,cD
%Q3_final: n,x,c_AF,c_TF,c_DF,L,S_total
    S=struct();
    if exist(xlsx,'file')
        try
            S.Q2_parts=readtable(xlsx,'Sheet','Q2_parts');
            S.Q2_final=readtable(xlsx,'Sheet','Q2_final');
            S.Q3_parts=readtable(xlsx,'Sheet','Q3_parts');
            S.Q3_half=readtable(xlsx,'Sheet','Q3_half');
            S.Q3_final=readtable(xlsx,'Sheet','Q3_final');
            return;
        catch ME
            warning('读取%s 失败（%s），改用示例基线。',xlsx,ME.message);
        end
    end
%示例基线：与前面Q2/Q3参数一致，可直接跑通 
    S.Q2_parts=cell2table({...
        'P1',100,10,6,2;...
        'P2',100,10,10,2},...
        'VariableNames',{'name','n','x','c_buy','c_test'});
    S.Q2_final=table(100,10,8,6,10,10,56,...
        'VariableNames',{'n','x','c_AF','c_TF','c_DF','L','S_total'});
    names=arrayfun(@(i)sprintf('Part%d',i),1:8,'UniformOutput',false)';
    ncol=repmat(100,8,1);
    xcol=repmat(10,8,1);
    S.Q3_parts=table(names,ncol,xcol,...
        'VariableNames',{'name','n','x'});
    S.Q3_parts.c_buy=[2 8 12 2 8 12 8 12]';
    S.Q3_parts.c_test=[1 1  2 1 1  2 1  2]';
    S.Q3_half=cell2table({...
        'A',100,10,8,4,6;...
        'B',100,10,8,4,6;...
        'C',100,10,8,4,6},...
        'VariableNames',{'name','n','x','cA','cT','cD'});
    S.Q3_final=table(100,10,8,6,10,40,200,...
        'VariableNames',{'n','x','c_AF','c_TF','c_DF','L','S_total'});
end
function Q2=run_Q2_mc(S,prior,M,OUT_ROOT,XLSX,DPI,TOPK)
    outDir=fullfile(OUT_ROOT,'Q2_mean');
    if ~exist(outDir,'dir');mkdir(outDir);end
%后验：良率=1 - BetaPosterior
    p1=1 - betarnd(prior(1)+S.Q2_parts.x(1),prior(2)+S.Q2_parts.n(1)-S.Q2_parts.x(1),M,1);
    p2=1 - betarnd(prior(1)+S.Q2_parts.x(2),prior(2)+S.Q2_parts.n(2)-S.Q2_parts.x(2),M,1);
    yF=1 - betarnd(prior(1)+S.Q2_final.x(1),prior(2)+S.Q2_final.n(1)-S.Q2_final.x(1),M,1);
    costP=[S.Q2_parts.c_buy(1) S.Q2_parts.c_test(1) S.Q2_parts.c_buy(2) S.Q2_parts.c_test(2)];
    cAF=S.Q2_final.c_AF(1);cTF=S.Q2_final.c_TF(1);cDF=S.Q2_final.c_DF(1);
    L=S.Q2_final.L(1);S_total=S.Q2_final.S_total(1);
%与你Q2摘要一致的6种策略
    policies={ ...
        [1 1 0 1];[0 1 0 1];[1 0 0 1];...
        [1 0 0 0];[0 1 1 1];[1 1 1 1]};
    names={ ...
        '预检1=是，预检2=是，终检=否，拆=是';...
        '预检1=否，预检2=是，终检=否，拆=是';...
        '预检1=是，预检2=否，终检=否，拆=是';...
        '预检1=是，预检2=否，终检=否，拆=否';...
        '预检1=否，预检2=是，终检=是，拆=是';...
        '预检1=是，预检2=是，终检=是，拆=是'};
    K=numel(policies);
    mu=zeros(K,1);sd=zeros(K,1);p05=zeros(K,1);p95=zeros(K,1);
    for k=1:K
        bits=policies{k};
        prof=zeros(M,1);
        for m=1:M
            prof(m)=q2_profit(bits,p1(m),p2(m),yF(m),costP,cAF,cTF,cDF,L,S_total);
        end
        mu(k)=mean(prof);
        sd(k)=std(prof);
        p05(k)=prctile(prof,5);
        p95(k)=prctile(prof,95);
    end
    T=table(names,mu,sd,p05,p95,'VariableNames',{'策略','均值利润','Std','P05','P95'});
    writetable(T,XLSX,'Sheet','Q2_均值与分位');
    [~,ord]=sort(mu,'descend');keep=min(TOPK,K);
    f=figure('Color','w','Position',[80 80 1000 580]);
    bar(mu(ord(1:keep)));
    hold on;
    er=errorbar(1:keep,mu(ord(1:keep)),mu(ord(1:keep))-p05(ord(1:keep)),p95(ord(1:keep))-mu(ord(1:keep)));
    er.LineStyle='none';
    grid on;
    ylabel('利润(元/单)');
    title('Q4–Q2：策略利润均值与5%–95%区间');
    set(gca,'XTick',1:keep,'XTickLabel',names(ord(1:keep)),'XTickLabelRotation',20);
    saveFig_noToolbar(f,fullfile(outDir,'Q2_profit_mean_p05p95.png'),DPI);
    Q2=struct('table',T,'order',ord);
end
function val=q2_profit(bits,p1,p2,yF,costP,cAF,cTF,cDF,L,S_total)
    I1=bits(1);I2=bits(2);IF=bits(3);D=bits(4);
    cb1=costP(1);ct1=costP(2);
    cb2=costP(3);ct2=costP(4);
    C_ins=I1*((cb1+ct1)/max(p1,realmin))+I2*((cb2+ct2)/max(p2,realmin));
    C_un=(1-I1)*cb1+(1-I2)*cb2;
    q_ship=(I1*1+(1-I1)*p1) * (I2*1+(1-I2)*p2) * yF;
    if IF==1
        EN=1/max(q_ship,realmin);
        cost=C_ins+EN*(C_un+cAF+cTF)+(EN-1)*D*cDF;
        profit=S_total - cost;
    else
        base=C_ins+C_un+cAF;
        ratio=(1 - q_ship)/max(q_ship,realmin);
        crework=ratio * (cDF+cAF+cTF);
        cret=ratio * L;
        cost=base+crework+cret;
        profit=S_total - cost;
    end
    val=profit;
end
function Q3=run_Q3_mc(S,prior,M,OUT_ROOT,XLSX,DPI,TOPK)
    outDir=fullfile(OUT_ROOT,'Q3_mean');
    if ~exist(outDir,'dir');mkdir(outDir);end
%后验采样：8零件、3半段、成段 
%逐列采样，避免 betarnd 尺寸不匹配
    p_part=zeros(M,8);
    for j=1:8
        a=prior(1)+S.Q3_parts.x(j);
        b=prior(2)+S.Q3_parts.n(j) - S.Q3_parts.x(j);
        p_part(:,j)=1 - betarnd(a,b,M,1);
    end
    y_half=zeros(M,3);
    for j=1:3
        a=prior(1)+S.Q3_half.x(j);
        b=prior(2)+S.Q3_half.n(j) - S.Q3_half.x(j);
        y_half(:,j)=1 - betarnd(a,b,M,1);
    end
    a=prior(1)+S.Q3_final.x(1);
    b=prior(2)+S.Q3_final.n(1) - S.Q3_final.x(1);
    y_final=1 - betarnd(a,b,M,1);
    cb=S.Q3_parts.c_buy';ct=S.Q3_parts.c_test';
    cA=S.Q3_half.cA';cT=S.Q3_half.cT';cD=S.Q3_half.cD';
    cAF=S.Q3_final.c_AF(1);cTF=S.Q3_final.c_TF(1);cDF=S.Q3_final.c_DF(1);
    L=S.Q3_final.L(1);S_total=S.Q3_final.S_total(1);
    nPol=4096;
    mu=zeros(nPol,1);
    p05=zeros(nPol,1);
    p95=zeros(nPol,1);
    pol_str=cell(nPol,1);
    for k=0:nPol-1
        bits=bitget(uint16(k),12:-1:1);
        pol_str{k+1}=sprintf('%s | 半检=%s | 半拆=%s | 成检=%s | 成拆=%s',...
            sprintf('%d',bits(1:8)),yn(bits(9)),yn(bits(10)),yn(bits(11)),yn(bits(12)));
        prof=zeros(M,1);
        for m=1:M
            P=struct();
            P.p_part=p_part(m,:);
            P.y_half=y_half(m,:);
            P.y_final= y_final(m);
            P.c_buy=cb;P.c_test_part=ct;
            P.c_asm_half=cA;P.c_test_half=cT;P.c_dis_half=cD;
            P.c_asm_final=cAF;P.c_test_final=cTF;P.c_dis_final=cDF;
            P.price=S_total;P.exch_loss=L;P.aftersale_rework_with_test=true;
            out=eval_policy_bits_g(bits,P);
            prof(m)=out.profit_per_customer;
        end
        mu(k+1)=mean(prof);
        p05(k+1)=prctile(prof,5);
        p95(k+1)=prctile(prof,95);
    end
    T=table((1:nPol)',pol_str,mu,p05,p95,...
        'VariableNames',{'index','policy','Profit_mean','Profit_p05','Profit_p95'});
    Tsort=sortrows(T,'Profit_mean','descend');
    writetable(Tsort,XLSX,'Sheet','Q3_4096_均值与分位');
    K=min(TOPK,height(Tsort));
    f=figure('Color','w','Position',[80 80 1200 600]);
    bar(Tsort.Profit_mean(1:K));
    hold on;
    er=errorbar(1:K,Tsort.Profit_mean(1:K),...
        Tsort.Profit_mean(1:K)-Tsort.Profit_p05(1:K),...
        Tsort.Profit_p95(1:K)-Tsort.Profit_mean(1:K));
    er.LineStyle='none';
    grid on;
    ylabel('利润(元/单)');
    title('Q4–Q3：总体TopK（利润均值与5%–95%）');
    set(gca,'XTick',1:K,'XTickLabel',Tsort.policy(1:K),'XTickLabelRotation',20);
    saveFig_noToolbar(f,fullfile(outDir,'Q3_TopK_mean_p05p95.png'),DPI);
    Q3=struct('table',Tsort);
end
%三段严谨口径
function out=eval_policy_bits_g(bits,P)
    idxA=[1 2 3];idxB=[4 5 6];idxC=[7 8];
    insp_part=logical(bits(1:8));
    half_insp=bits(9)==1;
    half_dis=bits(10)==1;
    final_insp= bits(11)==1;
    final_dis=bits(12)==1;
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
                cost=Nf * ((H1.C_good+H2.C_good+H3.C_good)+cAF+cTF);
                comp_base=cost;comp_rework=0;comp_fdis=0;
            end
            q_ship=1;N_final=Nf;
        else
            s_final=sH1 * sH2 * sH3 * yF;
            Nf=1/max(s_final,eps);
            cost_once=(H1.C_attempt+H2.C_attempt+H3.C_attempt+cAF+cTF);
            if final_dis
                comp_base=cost_once;
                comp_rework=(Nf-1)*cost_once;
                comp_fdis=(Nf-1)*cDF;
            else
                comp_base=Nf*cost_once;
                comp_rework=0;comp_fdis=0;
            end
            cost=comp_base+comp_rework+comp_fdis;
            q_ship=1;N_final=Nf;
        end
        comp_return=0;
        cost_per_customer=cost;
        profit_per_customer=PRICE - cost_per_customer;
    else
        if half_insp
            base_cost=(H1.C_good+H2.C_good+H3.C_good+cAF);
            if isfield(P,'aftersale_rework_with_test') && P.aftersale_rework_with_test
                if final_dis
                    E_rework=(cDF+cAF+cTF) / max(yF,eps);
                else
                    E_rework=((H1.C_good+H2.C_good+H3.C_good)+cAF+cTF) / max(yF,eps);
                end
                comp_base=base_cost;
                comp_rework=(1 - yF) * E_rework;
                comp_return=(1 - yF) * LOSS;
                comp_fdis=0;
                cost_per_customer=comp_base+comp_rework+comp_return;
                profit_per_customer=PRICE - cost_per_customer;
                q_ship=yF;N_final=1+(1 - yF);
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
            s_ship=max(sH1*sH2*sH3*yF,eps);
            N_ship=1/s_ship;
            cost_once=(H1.C_attempt+H2.C_attempt+H3.C_attempt+cAF);
            comp_base=cost_once;comp_rework=(N_ship-1)*cost_once;comp_return=(N_ship-1)*LOSS;
            comp_fdis=0;
            cost_per_customer=comp_base+comp_rework+comp_return;
            profit_per_customer=PRICE-cost_per_customer;
            q_ship=s_ship;N_final=N_ship;
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
    out.comp_return_loss= comp_return;
    out.comp_final_dis=comp_fdis;
end
function [H,sH,N_half]=half_block_g(idx,hID,insp_part,half_insp,half_dis,...
    p,cb,ct,yH,cAH,cTH,cDH,eps)
    if numel(yH)<3;yH=repmat(yH(1),1,3);end
    if numel(cAH)<3;cAH=repmat(cAH(1),1,3);end
    if numel(cTH)<3;cTH=repmat(cTH(1),1,3);end
    if numel(cDH)<3;cDH=repmat(cDH(1),1,3);end
    r=ones(1,numel(idx));
    bad=~insp_part(idx);
    r(bad)=p(idx(bad));
    sH=prod(r) * yH(hID);
    sH=max(sH,eps);
    N_half=1/sH;
    C_ins_parts_per_attempt=sum( ((cb(idx)+ct(idx))./max(p(idx),eps)) .* insp_part(idx) );
    C_unins_per_attempt=sum( cb(idx(bad)) );
    C_attempt_no_test=C_ins_parts_per_attempt+C_unins_per_attempt+cAH(hID);
    C_attempt_with_test=C_attempt_no_test+cTH(hID);
    H=struct();
    if half_insp
        F=N_half - 1;
        if half_dis
            C_ins_once=sum( ((cb(idx)+ct(idx))./max(p(idx),eps)) .* insp_part(idx) );
            C_unins_each=sum( cb(idx(bad)) );
            H.C_good=C_ins_once+N_half*(C_unins_each+cAH(hID)+cTH(hID))+F*cDH(hID);
        else
            H.C_good=N_half * (C_ins_parts_per_attempt+C_unins_per_attempt+cAH(hID)+cTH(hID));
        end
        H.C_attempt=C_attempt_with_test;
    else
        H.C_good=C_attempt_no_test;
        H.C_attempt=C_attempt_no_test;
        N_half=1;
    end
end
function y=yn(b)
    if b~=0
        y='是';
    else
        y='否';
    end
end
function fname=chooseChineseFont()
    cands={'Microsoft YaHei','SimHei','PingFang SC','Heiti SC','Songti SC',...
             'Noto Sans CJK SC','Source Han Sans SC','Hiragino Sans GB'};
    avail=listfonts;
    fname=get(groot,'defaultAxesFontName');
    for i=1:numel(cands)
        if any(strcmp(avail,cands{i}))
            fname=cands{i};
            return;
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
        [p,~,~]=fileparts(filename);
        if ~exist(p,'dir');mkdir(p);end
        print(f,filename,'-dpng',sprintf('-r%d',dpi));
    end
end
function s=bits2str_g(bits)
s=sprintf('%s | 半检=%s | 半拆=%s | 成检=%s | 成拆=%s',...
    sprintf('%d',bits(1:8)),yn(bits(9)),yn(bits(10)),yn(bits(11)),yn(bits(12)));
end