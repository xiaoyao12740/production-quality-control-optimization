%main4_2.m-问题4：Bayes+风险决策（Q2+Q3）
clc;clear;close all;tic;
%运行开关 
RUN_Q2=true;
RUN_Q3=true;
%输出 
OUT_ROOT=fullfile(pwd,'q4_output');
if ~exist(OUT_ROOT,'dir'),mkdir(OUT_ROOT);end
XLSX=fullfile(OUT_ROOT,'q4_results.xlsx');
SAVE_DPI=240;TOPK=16;
%外观 
cnFont=chooseChineseFont();
set(groot,'defaultAxesFontName',cnFont,'defaultTextFontName',cnFont,...
          'defaultAxesFontSize',11,'defaultTextFontSize',10,'defaultFigureColor','w');
try set(groot,'defaultAxesToolbarVisible','off');catch,end
%Monte Carlo 后验设置 
M=2000;%后验样本条数（可调 1000~3000）
prior=[0.5,0.5];%Beta(1/2,1/2) Jeffreys先验
%读取抽样/参数（与 main4_1一致的Sheet约定）
S=load_sampling_or_default('q4_input.xlsx');%兜底示例
%生成后验样本 
samples=posterior_from_counts(S,prior,M);
%Q2：16策略（I1/I2/IF/D） 
if RUN_Q2
    outDir=fullfile(OUT_ROOT,'Q2_risk');
    if ~exist(outDir,'dir'),mkdir(outDir);end
%16种策略位
    pol16=dec2bin(0:15,4)-'0';
    names=arrayfun(@(i) strategyNameQ2(pol16(i,:)),1:16,'uni',0);
%成本参数解包
    costP=[S.Q2_parts.c_buy(1) S.Q2_parts.c_test(1) S.Q2_parts.c_buy(2) S.Q2_parts.c_test(2)];
    cAF=S.Q2_final.c_AF(1);cTF=S.Q2_final.c_TF(1);cDF=S.Q2_final.c_DF(1);
    L=S.Q2_final.L(1);S_total=S.Q2_final.S_total(1);
    mu=zeros(16,1);sd=mu;p05=mu;p95=mu;var5=mu;cvar5=mu;prob_best=mu;
    PROF=zeros(M,16);%用于最优概率/EVPI
%遍历16策略
    for k=1:16
        bits=pol16(k,:);
        prof_k=q2_profit_vec(bits,samples.p1,samples.p2,samples.yF,...
                               costP,cAF,cTF,cDF,L,S_total);%Mx1
        [mu(k),sd(k),p05(k),p95(k),var5(k),cvar5(k)]=risk_summary(prof_k);
        PROF(:,k)=prof_k;
    end
%后验最优概率
    [~,imax]=max(PROF,[],2);
    for k=1:16
        prob_best(k)=mean(imax==k);
    end
%EVPI（粗）：先验下"知道真值再选"的期望 - "按Bayes均值选"的期望
    evpi=evpi_estimate(PROF);
    Tq2=table(names(:),mu,sd,p05,p95,var5,cvar5,prob_best,...
        'VariableNames',{'策略','均值','Std','P05','P95','VaR5','CVaR5','后验最优概率'});
    writetable(Tq2,XLSX,'Sheet','Q2_风险统计','WriteMode','overwritesheet');
%可视化：按不同准则排序（示例：按均值/按CVaR5）
    [~,ord_mu]=sort(mu,'descend');keep=min(TOPK,16);
    f1=figure('Color','w','Position',[80 80 1000 580]);
    bar(mu(ord_mu(1:keep)));grid on;ylabel('利润(元/单)');
    title(sprintf('Q4–Q2：均值Top%d（EVPI=%.2f）',keep,evpi));
    set(gca,'XTick',1:keep,'XTickLabel',names(ord_mu(1:keep)),'XTickLabelRotation',20);
    saveFig_noToolbar(f1,fullfile(outDir,'Q2_Top_by_mean.png'),SAVE_DPI);
    [~,ord_cvar]=sort(cvar5,'descend');
    f2=figure('Color','w','Position',[80 80 1000 580]);
    bar(cvar5(ord_cvar(1:keep)));grid on;ylabel('CVaR(5%) 利润');
    title('Q4–Q2：CVaR(5%) Top');
    set(gca,'XTick',1:keep,'XTickLabel',names(ord_cvar(1:keep)),'XTickLabelRotation',20);
    saveFig_noToolbar(f2,fullfile(outDir,'Q2_Top_by_CVaR5.png'),SAVE_DPI);
end
%Q3：4096策略 
if RUN_Q3
    outDir=fullfile(OUT_ROOT,'Q3_risk');
    if ~exist(outDir,'dir'),mkdir(outDir);end
    nPol=4096;
    mu=zeros(nPol,1);sd=mu;p05=mu;p95=mu;var5=mu;cvar5=mu;prob_best=zeros(nPol,1);
%预生成 4096 位图矩阵（12位）
    bitsMat=logical(dec2bin(0:nPol-1,12)-'0');
%为了节省内存，不保存 PROF 全矩阵，分块/并行统计
    chunk=512;%每块策略数（可按内存调）
    nBlk=ceil(nPol/chunk);
    postBestCount=zeros(nPol,1);%统计"本策略在某样本上为最优"的次数
    bestMeanSoFar=-inf;%仅用于打印信息
%预分配
mu=zeros(nPol,1);sd=mu;p05=mu;p95=mu;var5=mu;cvar5=mu;
%用于"全局最优策略（逐样本）"的流式聚合
bestVal=-inf(M,1);%每个样本当前最大利润
bestIdx=zeros(M,1);%其对应的全局策略索引（1..nPol）
for blk=1:nBlk%外层顺序
    i1=(blk-1)*chunk+1;
    i2=min(blk*chunk,nPol);
    K=i2 - i1+1;
    PROF_blk=zeros(M,K);%本块 M×K 的利润矩阵
%内层并行：每个策略各自算 M 条样本的利润向量 
    parfor kk=1:K%内层并行
        i=i1+kk - 1;%全局策略索引
        b=bitsMat(i,:);
        prof_k=zeros(M,1);
%逐样本（可按你原逻辑）
        for m=1:M
            P=struct();
            P.p_part=samples.p_part(m,:);
            P.y_half=samples.y_half(m,:);
            P.y_final=samples.y_final(m);
            P.c_buy=samples.cb;P.c_test_part=samples.ct;
            P.c_asm_half=samples.cA;P.c_test_half=samples.cT;P.c_dis_half=samples.cD;
            P.c_asm_final=samples.cAF;P.c_test_final=samples.cTF;P.c_dis_final=samples.cDF;
            P.price=samples.S_total;P.exch_loss=samples.L;P.aftersale_rework_with_test=true;
            out=eval_policy_bits_g(b,P);
            prof_k(m)=out.profit_per_customer;
        end
        PROF_blk(:,kk)=prof_k;
    end
%本块统计（均值/分位/VaR/CVaR） 
    mu_b=mean(PROF_blk,1).';%K×1
    sd_b=std(PROF_blk,0,1).';
    p05_b=prctile(PROF_blk,5).';
    p95_b=prctile(PROF_blk,95).';
    var5_b=p05_b;
    cvar5_b=zeros(K,1);
    kcut=max(1,floor(0.05*M));
    for kk=1:K
        xx=sort(PROF_blk(:,kk));
        cvar5_b(kk)=mean(xx(1:kcut));
    end
%回写到全局数组（现在是普通 for，切片赋值完全合法） 
    mu(i1:i2)=mu_b;sd(i1:i2)=sd_b;p05(i1:i2)=p05_b;p95(i1:i2)=p95_b;
    var5(i1:i2)=var5_b;cvar5(i1:i2)=cvar5_b;
%更新"全局最优策略（每个样本）"
    [locMax,imax_blk]=max(PROF_blk,[],2);%本块内最优
    upd=locMax > bestVal;%与当前全局最优比较
    bestVal(upd)=locMax(upd);
    bestIdx(upd)=i1 - 1+imax_blk(upd);%变成全局策略编号
end
%汇总"后验最优概率"
postBestCount=accumarray(bestIdx,1,[nPol,1]);
prob_best=postBestCount / M;
%排序与导出
    [~,ord]=sort(mu,'descend');
    keep=min(TOPK,numel(ord));
    pol_str=arrayfun(@(i) bits2str_g(bitsMat(i,:)),1:nPol,'uni',0)';
    Tq3=table((1:nPol)',pol_str,mu,sd,p05,p95,var5,cvar5,prob_best,...
        'VariableNames',{'index','policy','均值','Std','P05','P95','VaR5','CVaR5','后验最优概率'});
    writetable(Tq3,XLSX,'Sheet','Q3_4096_风险统计','WriteMode','overwritesheet');
%可视化：TopK（按均值）
    f=figure('Color','w','Position',[80 80 1200 580]);
    bar(mu(ord(1:keep)));grid on;ylabel('利润(元/单)');
    title(sprintf('Q4–Q3：均值Top%d',keep));
    set(gca,'XTick',1:keep,'XTickLabel',pol_str(ord(1:keep)),'XTickLabelRotation',20);
    saveFig_noToolbar(f,fullfile(outDir,'Q3_Top_by_mean.png'),SAVE_DPI);
%可视化：TopK（按CVaR5）
    [~,ordc]=sort(cvar5,'descend');
    f2=figure('Color','w','Position',[80 80 1200 580]);
    bar(cvar5(ordc(1:keep)));grid on;ylabel('CVaR(5%) 利润');
    title(sprintf('Q4–Q3：CVaR(5%%) Top%d',keep));
    set(gca,'XTick',1:keep,'XTickLabel',pol_str(ordc(1:keep)),'XTickLabelRotation',20);
    saveFig_noToolbar(f2,fullfile(outDir,'Q3_Top_by_CVaR5.png'),SAVE_DPI);
end
fprintf('\n main4_2 完成。输出：%s\nExcel：%s\n耗时：%.2f s\n',OUT_ROOT,XLSX,toc);
%= 依赖函数（可移入 utils_q4.m）=
function S=load_sampling_or_default(xlsx)
    S=struct();
    if exist(xlsx,'file')
        try
            S.Q2_parts=readtable(xlsx,'Sheet','Q2_parts');
            S.Q2_final=readtable(xlsx,'Sheet','Q2_final');
            S.Q3_parts=readtable(xlsx,'Sheet','Q3_parts');
            S.Q3_half=readtable(xlsx,'Sheet','Q3_half');
            S.Q3_final=readtable(xlsx,'Sheet','Q3_final');
            return;
        catch
            warning('读取%s 失败，改用示例基线。',xlsx);
        end
    end
%兜底示例（与你 main4_1 相同）
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

function samples=posterior_from_counts(S,prior,M)
%Q2
    samples.p1=1 - betarnd(prior(1)+S.Q2_parts.x(1),prior(2)+S.Q2_parts.n(1)-S.Q2_parts.x(1),M,1);
    samples.p2=1 - betarnd(prior(1)+S.Q2_parts.x(2),prior(2)+S.Q2_parts.n(2)-S.Q2_parts.x(2),M,1);
    samples.yF=1 - betarnd(prior(1)+S.Q2_final.x(1),prior(2)+S.Q2_final.n(1)-S.Q2_final.x(1),M,1);
    samples.cb=S.Q3_parts.c_buy';samples.ct=S.Q3_parts.c_test';
    samples.cA=S.Q3_half.cA';samples.cT=S.Q3_half.cT';samples.cD=S.Q3_half.cD';
    samples.cAF=S.Q3_final.c_AF(1);samples.cTF=S.Q3_final.c_TF(1);samples.cDF=S.Q3_final.c_DF(1);
    samples.L=S.Q3_final.L(1);samples.S_total=S.Q3_final.S_total(1);
%Q3
    Mx=M;
    p_part=zeros(Mx,8);
    for j=1:8
        a=prior(1)+S.Q3_parts.x(j);
        b=prior(2)+S.Q3_parts.n(j)-S.Q3_parts.x(j);
        p_part(:,j)=1 - betarnd(a,b,Mx,1);
    end
    y_half=zeros(Mx,3);
    for j=1:3
        a=prior(1)+S.Q3_half.x(j);
        b=prior(2)+S.Q3_half.n(j)-S.Q3_half.x(j);
        y_half(:,j)=1 - betarnd(a,b,Mx,1);
    end
    a=prior(1)+S.Q3_final.x(1);
    b=prior(2)+S.Q3_final.n(1)-S.Q3_final.x(1);
    y_final=1 - betarnd(a,b,Mx,1);
    samples.p_part=p_part;
    samples.y_half=y_half;
    samples.y_final=y_final;
end
function prof=q2_profit_vec(bits,p1,p2,yF,costP,cAF,cTF,cDF,L,S_total)
    I1=bits(1);I2=bits(2);IF=bits(3);D=bits(4);
    cb1=costP(1);ct1=costP(2);cb2=costP(3);ct2=costP(4);
    C_ins=I1*((cb1+ct1)./max(p1,realmin))+I2*((cb2+ct2)./max(p2,realmin));
    C_un=(1-I1)*cb1+(1-I2)*cb2;
    q_ship=(I1+(1-I1).*p1) .* (I2+(1-I2).*p2) .* yF;
    if IF==1
        EN=1./max(q_ship,realmin);
        cost=C_ins+EN.*(C_un+cAF+cTF)+(EN-1).*D.*cDF;
        prof=S_total - cost;
    else
        base=C_ins+C_un+cAF;
        ratio=(1 - q_ship)./max(q_ship,realmin);
        crework=ratio.*(cDF+cAF+cTF);
        cret=ratio.*L;
        cost=base+crework+cret;
        prof=S_total - cost;
    end
end
function [mu,sd,p05,p95,VaR5,CVaR5]=risk_summary(x)
    mu=mean(x);sd=std(x);
    p05=prctile(x,5);p95=prctile(x,95);
    VaR5=p05;
%CVaR5：下5%均值
    xsort=sort(x);k=max(1,floor(0.05*numel(xsort)));
    CVaR5=mean(xsort(1:k));
end
function ev=evpi_estimate(PROF)
%PROF: M x K
%EVPI ≈ E_m[ max_k PROF(m,k) ] - max_k E_m[ PROF(m,k) ]
    ev=mean(max(PROF,[],2)) - max(mean(PROF,1));
end
function s=strategyNameQ2(bits)
    yn=@(b) ternary(b,'是','否');
    s=sprintf('预检1=%s，预检2=%s，终检=%s，拆=%s',yn(bits(1)),yn(bits(2)),yn(bits(3)),yn(bits(4)));
end
function y=ternary(c,a,b),if c,y=a;else,y=b;end,end
%复用Q3三段实现：eval_policy_bits_g / bits2str_g / chooseChineseFont / saveFig_noToolbar===
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
end
function s=bits2str_g(bits)
s=sprintf('%s | 半检=%s | 半拆=%s | 成检=%s | 成拆=%s',...
    sprintf('%d',bits(1:8)),yn(bits(9)),yn(bits(10)),yn(bits(11)),yn(bits(12)));
end