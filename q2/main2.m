%main2
clc;clear;close all;
%可调参数
OUT_ROOT=fullfile(pwd,'q2_output');%输出根目录
XLSX_FILE=fullfile(OUT_ROOT,'q2_results.xlsx');
SENS_DELTA=0.10;%灵敏度：±10%
TOP_N=6;%Top显示个数
SAVE_DPI=220;%PNG分辨率
SHOW_VALUE_LABELS=true;%柱顶显示数值
SAVE_TEXT_SUMMARY=true;%同步把命令行要点写到txt
if ~exist(OUT_ROOT,'dir'),mkdir(OUT_ROOT);end
%中文字体（只影响图）
cnFont=chooseChineseFont();
set(groot,'defaultAxesFontName',cnFont,...
    'defaultTextFontName',cnFont,...
    'defaultAxesFontSize',11,...
    'defaultTextFontSize',10,...
    'defaultFigureColor','w');
%表1：六种情形（按题面数据）
%列顺序: [p1 c_b1 c_t1, p2 c_b2 c_t2, p_f, c_a, c_tf, s, L, c_d]
cases=[...
    0.10 4 2,0.10 18 3,0.10,6,3,56,6,5;%情形1
    0.20 4 2,0.20 18 3,0.20,6,3,56,6,5;%情形2
    0.10 4 2,0.10 18 3,0.10,6,3,56,30,5;%情形3
    0.20 4 1,0.20 18 1,0.20,6,2,56,30,5;%情形4
    0.10 4 8,0.20 18 1,0.10,6,2,56,10,5;%情形5
    0.05 4 2,0.05 18 3,0.05,6,3,56,10,40];%情形6
case_names=arrayfun(@(k)sprintf('情形%d',k),1:size(cases,1),'uni',0);
%六种固定配色（便于一眼识别情形）
caseColors=[...
    0.22 0.49 0.72;%蓝
    0.30 0.69 0.29;%绿
    0.87 0.39 0.14;%橙
    0.60 0.31 0.64;%紫
    0.89 0.10 0.11;%红
    0.90 0.90 0.00];%黄
%穷举策略空间 I1,I2,IF,D ∈ {0,1}
%采用二进制枚举，列顺序固定为 [I1 I2 IF D]
policies=dec2bin(0:15,4)-'0';%16x4
nPol=size(policies,1);
%容器
all_results=cell(size(cases,1),1);
top_summary_rows=[];
%主循环
for k=1:size(cases,1)
    pars=cases(k,:);
    col=caseColors(k,:);
    casename=case_names{k};
    %按情形建立文件夹（用于放置图和TXT）
    caseDir=fullfile(OUT_ROOT,casename);
    if ~exist(caseDir,'dir'),mkdir(caseDir);end
    %计算所有策略并按利润降序
    [name_long,name_short,q_arr,cost_arr,prof_arr,en_arr,ret_arr,policies_sorted]=...
        evaluate_all_policies(pars,policies);
    txtfile=fullfile(caseDir,[casename '_summary.txt']);
    if SAVE_TEXT_SUMMARY,fid=fopen(txtfile,'w');else,fid=-1;end
    pf=@(varargin)fprintf(varargin{:});%控制台
    wf=@(varargin)ifwrite(fid,varargin{:});%文本
    pf('\n==== %s ====\n',casename);
    wf('==== %s ====\n',casename);
    pf('最优策略：%s\n',name_long{1});wf('最优策略：%s\n',name_long{1});
    pf('  q=%.4f, E[N]=%.2f, E_return=%.2f, 成本=%.2f, 利润=%.2f\n',...
        q_arr(1),en_arr(1),ret_arr(1),cost_arr(1),prof_arr(1));
    wf('  q=%.4f, E[N]=%.2f, E_return=%.2f, 成本=%.2f, 利润=%.2f\n',...
        q_arr(1),en_arr(1),ret_arr(1),cost_arr(1),prof_arr(1));
    topN=min(TOP_N,nPol);
    pf('\n Top %d（按利润降序）\n',topN);
    wf('\n Top %d（按利润降序）\n',topN);
    for r=1:topN
        pf('%2d) %-52s | 利润=%7.2f | 成本=%7.2f | q=%6.4f | E[N]=%5.2f | 退换=%5.2f\n',...
            r,name_long{r},prof_arr(r),cost_arr(r),q_arr(r),en_arr(r),ret_arr(r));
        wf('%2d) %-52s | 利润=%7.2f | 成本=%7.2f | q=%6.4f | E[N]=%5.2f | 退换=%5.2f\n',...
            r,name_long{r},prof_arr(r),cost_arr(r),q_arr(r),en_arr(r),ret_arr(r));
    end
    %边际影响：翻转I1/I2/IF/D
    best_bits=policies_sorted(1,:);
    [deltaBitTbl]=bit_marginal_effect(best_bits,policies_sorted,prof_arr,name_short);
    pf('\n边际影响（基于最优，逐个翻转I1/I2/IF/D的Δ利润）：\n');wf('\n边际影响（逐个翻转I1/I2/IF/D）：\n');
    for i=1:height(deltaBitTbl)
        pf('  翻转%-3s：Δ利润=%+.2f → %s\n',deltaBitTbl.Bit{i},deltaBitTbl.DeltaProfit(i),deltaBitTbl.TargetName{i});
        wf('  翻转%-3s：Δ利润=%+.2f → %s\n',deltaBitTbl.Bit{i},deltaBitTbl.DeltaProfit(i),deltaBitTbl.TargetName{i});
    end
    %画图
    %图1利润柱状
    fP=figure('Name',[casename '：利润'],'Color','w','Position',[100 100 1100 500]);
    axP=axes(fP);bar(axP,prof_arr,'FaceColor',col,'EdgeColor','none');
    title([casename '：单件期望利润（元/件）']);ylabel('利润（元/件）');
    grid on;xlim([0 nPol+1]);set(axP,'XTick',1:nPol,'XTickLabel',name_short,'XTickLabelRotation',45);
    if SHOW_VALUE_LABELS,addBarValueLabels(axP,1:nPol,prof_arr,'%.2f');end
    saveFig(fP,fullfile(caseDir,'01_利润_rank.png'),SAVE_DPI);
    %图2成本柱状
    fC=figure('Name',[casename '：成本'],'Color','w','Position',[100 100 1100 500]);
    axC=axes(fC);bar(axC,cost_arr,'FaceColor',lighten(col,0.15),'EdgeColor','none');
    title([casename '：单件期望成本（元/件）']);ylabel('成本（元/件）');
    grid on;xlim([0 nPol+1]);set(axC,'XTick',1:nPol,'XTickLabel',name_short,'XTickLabelRotation',45);
    if SHOW_VALUE_LABELS,addBarValueLabels(axC,1:nPol,cost_arr,'%.2f');end
    saveFig(fC,fullfile(caseDir,'02_成本_rank.png'),SAVE_DPI);
    %图3散点：利润 vs q（直接用q_arr）
    fS=figure('Name',[casename '：利润 vs q'],'Color','w','Position',[100 100 900 500]);
    scatter(q_arr,prof_arr,50,col,'filled');grid on;
    title([casename '：利润 vs 成功率q']);xlabel('q');ylabel('利润（元/件）');
    saveFig(fS,fullfile(caseDir,'03_利润_vs_q_scatter.png'),SAVE_DPI);
    %图4帕累托前沿：利润最大化 & 成本最小化
    [paretoIdx]=pareto_front(cost_arr,prof_arr);%非支配解索引
    fPa=figure('Name',[casename '：帕累托前沿'],'Color','w','Position',[100 100 900 560]);
    scatter(cost_arr,prof_arr,36,[0.75 0.75 0.75],'filled');hold on;
    scatter(cost_arr(paretoIdx),prof_arr(paretoIdx),60,col,'filled');
    [~,ordc]=sort(cost_arr(paretoIdx),'ascend');
    plot(cost_arr(paretoIdx(ordc)),prof_arr(paretoIdx(ordc)),'-','Color',col,'LineWidth',1.2);
    grid on;xlabel('成本（元/件）');ylabel('利润（元/件）');title([casename '：利润-成本 帕累托前沿']);
    legend({'策略全集','非支配解','前沿连线'},'Location','best');
    saveFig(fPa,fullfile(caseDir,'04_帕累托_利润成本.png'),SAVE_DPI);
    %图5退换次数茎叶图
    fR=figure('Name',[casename '：退换次数'],'Color','w','Position',[100 100 1100 450]);
    stem(1:nPol,ret_arr,'filled','Color',col);grid on;
    title([casename '：每单期望退换次数E[return]']);ylabel('次数/单');xlim([0 nPol+1]);
    set(gca,'XTick',1:nPol,'XTickLabel',name_short,'XTickLabelRotation',45);
    saveFig(fR,fullfile(caseDir,'05_退换次数_stem.png'),SAVE_DPI);
    %灵敏度（±10%），并出"龙卷风图"
    sens=do_sensitivity(pars,policies,name_long{1},SENS_DELTA);
    %文本摘要
    [top3chg,top3name]=topK(abs(sens.delta_profit_plus),sens.param_names,3);
    pf('\n灵敏度摘要（±%d%）\n',round(100*SENS_DELTA));
    wf('\n灵敏度摘要（±%d%）\n',round(100*SENS_DELTA));
    pf('  +%d% 时最敏感Top3：%s（|Δ|=%.2f）, %s（%.2f）, %s（%.2f）\n',...
        round(100*SENS_DELTA),top3name{1},top3chg(1),top3name{2},top3chg(2),top3name{3},top3chg(3));
    wf('  +%d% 时最敏感Top3：%s（|Δ|=%.2f）, %s（%.2f）, %s（%.2f）\n',...
        round(100*SENS_DELTA),top3name{1},top3chg(1),top3name{2},top3chg(2),top3name{3},top3chg(3));
    pf('  最优策略更换次数：+%d%= %d次，-%d%= %d次\n',round(100*SENS_DELTA),sens.best_switch_count_plus,round(100*SENS_DELTA),sens.best_switch_count_minus);
    wf('  最优策略更换次数：+%d%= %d次，-%d%= %d次\n',round(100*SENS_DELTA),sens.best_switch_count_plus,round(100*SENS_DELTA),sens.best_switch_count_minus);
    %龙卷风（水平条形：左负右正）
    fT=figure('Name',[casename '：灵敏度-龙卷风'],'Color','w','Position',[100 100 1100 640]);
    [~,ordT]=sort(max(abs([sens.delta_profit_plus sens.delta_profit_minus]),[],2),'descend');
    showTop=min(12,numel(ordT));id=ordT(1:showTop);
    ylab=sens.param_names(id);
    subplot(1,2,1);
    barh(-abs(sens.delta_profit_minus(id)),'FaceColor',lighten(col,0.25),'EdgeColor','none');
    set(gca,'YDir','reverse','YTick',1:showTop,'YTickLabel',ylab);
    xlabel('利润变化（元）');title(['-' num2str(round(100*SENS_DELTA)) '%扰动']);grid on;
    subplot(1,2,2);
    barh(abs(sens.delta_profit_plus(id)),'FaceColor',col,'EdgeColor','none');
    set(gca,'YDir','reverse','YTick',1:showTop,'YTickLabel',ylab);
    xlabel('利润变化（元）');title(['+' num2str(round(100*SENS_DELTA)) '%扰动']);grid on;
    sgtitle([casename '：最优策略利润对参数的龙卷风敏感性'],'FontWeight','bold');
    saveFig(fT,fullfile(caseDir,'06_灵敏度_龙卷风.png'),SAVE_DPI);
    %写入xlsx（情形sheet+帕累托+灵敏度）
    T=table((1:nPol)',name_long,name_short,...
        policies_sorted(:,1),policies_sorted(:,2),policies_sorted(:,3),policies_sorted(:,4),...
        q_arr,en_arr,ret_arr,cost_arr,prof_arr,...
        'VariableNames',{'Rank','策略全称','策略简称','I1','I2','IF','D','q','E_N','E_return','E_cost','Profit'});
    writetable_safe(T,XLSX_FILE,casename);
    %帕累托sheet
    TP=T(paretoIdx,:);TP=sortrows(TP,{'E_cost','Profit'},{'ascend','descend'});
    writetable_safe(TP,XLSX_FILE,[casename '_帕累托']);
    %灵敏度sheet
    TSENS=table(sens.param_names(:),sens.delta_profit_plus(:),sens.delta_profit_minus(:),...
        'VariableNames',{'参数','Δ利润(+10%)','Δ利润(-10%)'});
    writetable_safe(TSENS,XLSX_FILE,[casename '_灵敏度']);
    %汇总Top sheet数据
    r2=min(2,nPol);
    bestRow={casename,name_long{1},prof_arr(1),cost_arr(1),q_arr(1),en_arr(1),ret_arr(1)};
    second={casename,name_long{r2},prof_arr(r2),cost_arr(r2),q_arr(r2),en_arr(r2),ret_arr(r2)};
    gap=prof_arr(1)-prof_arr(r2);
    top_summary_rows=[top_summary_rows;bestRow;second;{'','',gap,NaN,NaN,NaN,NaN}];%#ok<AGROW>
    if fid>0,fclose(fid);end
    %保存结构到内存
    all_results{k}=struct('case_name',casename,'policies',policies_sorted,...
        'name_long',{name_long},'name_short',{name_short},...
        'q',q_arr,'E_cost',cost_arr,'Profit',prof_arr,...
        'E_N',en_arr,'E_return',ret_arr);
end
%写汇总Top
TS=cell2table(top_summary_rows,...
    'VariableNames',{'情形','策略','利润','成本','q','E_N','E_return'});
writetable_safe(TS,XLSX_FILE,'汇总Top');
%放入工作区
assignin('base','all_results_B2',all_results);
fprintf('\n已完成：所有图与txt已按"情形"写入 %s\n',OUT_ROOT);
fprintf('Excel → %s\n',XLSX_FILE);
%函数区
function ifwrite(fid,varargin)
if fid>0,fprintf(fid,varargin{:});end
end
function [name_long,name_short,q_arr,cost_arr,prof_arr,en_arr,ret_arr,policies_sorted]=...
    evaluate_all_policies(pars,policies)
%解包
p1=pars(1);c_b1=pars(2);c_t1=pars(3);
p2=pars(4);c_b2=pars(5);c_t2=pars(6);
p_f=pars(7);c_a=pars(8);c_tf=pars(9);
s=pars(10);L=pars(11);c_d=pars(12);
nPol=size(policies,1);
name_long=cell(nPol,1);name_short=cell(nPol,1);
q_arr=zeros(nPol,1);cost_arr=zeros(nPol,1);prof_arr=zeros(nPol,1);
en_arr=zeros(nPol,1);ret_arr=zeros(nPol,1);
for j=1:nPol
    I1=policies(j,1);I2=policies(j,2);IF=policies(j,3);D=policies(j,4);
    %一次成功概率
    t1=I1+(1-I1)*(1-p1);
    t2=I2+(1-I2)*(1-p2);
    q=t1*t2*(1-p_f);
    %每次尝试通用成本（不含固定/失败）
    C_try_base=(1-I1)*c_b1+(1-I2)*c_b2+c_a+IF*c_tf;
    if q<=0
        E_cost=inf;Profit=-inf;E_try=inf;E_return=inf;
    else
        if D==1
            %拆解回用：预检已知合格件只准备一次（固定成本）
            C_fixed=I1*(c_b1+c_t1)/(1-p1)+I2*(c_b2+c_t2)/(1-p2);
            C_fail=(1-IF)*L+c_d;
            E_cost=C_fixed+(1/q)*C_try_base+((1-q)/q)*C_fail;
        else
            %不回用：每次尝试都要重新准备已知合格件
            C_try=C_try_base+I1*(c_b1+c_t1)/(1-p1)+I2*(c_b2+c_t2)/(1-p2);
            C_fail=(1-IF)*L;
            E_cost=(1/q)*C_try+((1-q)/q)*C_fail;
        end
        Profit=s-E_cost;
        E_try=1/q;
        E_return=(1-IF)*(1-q)/q;
    end
    q_arr(j)=q;cost_arr(j)=round(E_cost,4);prof_arr(j)=round(Profit,4);
    en_arr(j)=round(E_try,4);ret_arr(j)=round(E_return,4);
    [nmLong,nmShort]=strategyNameCN(I1,I2,IF,D);
    name_long{j}=nmLong;name_short{j}=nmShort;
end
[~,ord]=sort(prof_arr,'descend');
name_long=name_long(ord);name_short=name_short(ord);
q_arr=q_arr(ord);cost_arr=cost_arr(ord);prof_arr=prof_arr(ord);
en_arr=en_arr(ord);ret_arr=ret_arr(ord);
policies_sorted=policies(ord,:);
end
function sens=do_sensitivity(pars,policies,~,delta)
param_names={'p1','c_b1','c_t1','p2','c_b2','c_t2','p_f','c_a','c_tf','s','L','c_d'};
base=pars(:)';
[nmL,~,~,~,prof_base,~,~,~]=evaluate_all_policies(base,policies);
best_base_name=nmL{1};
prof_best_base=prof_base(1);
dp_plus=zeros(numel(base),1);
dp_minus=zeros(numel(base),1);
sw_plus=0;sw_minus=0;
for i=1:numel(base)
    pars_plus=base;pars_minus=base;
    if ismember(i,[1,4,7])%概率截断
        pars_plus(i)=min(max(pars_plus(i)*(1+delta),1e-6),0.999999);
        pars_minus(i)=min(max(pars_minus(i)*(1-delta),1e-6),0.999999);
    else
        pars_plus(i)=max(pars_plus(i)*(1+delta),0);
        pars_minus(i)=max(pars_minus(i)*(1-delta),0);
    end
    [nm2,~,~,~,prof2,~,~,~]=evaluate_all_policies(pars_plus,policies);
    dp_plus(i)=prof2(1)-prof_best_base;
    if ~strcmp(nm2{1},best_base_name),sw_plus=sw_plus+1;end
    [nm3,~,~,~,prof3,~,~,~]=evaluate_all_policies(pars_minus,policies);
    dp_minus(i)=prof3(1)-prof_best_base;
    if ~strcmp(nm3{1},best_base_name),sw_minus=sw_minus+1;end
end
sens=struct('param_names',{param_names},...
    'delta_profit_plus',dp_plus,...
    'delta_profit_minus',dp_minus,...
    'best_switch_count_plus',sw_plus,...
    'best_switch_count_minus',sw_minus,...
    'best_base_name',best_base_name,...
    'prof_best_base',prof_best_base);
end
function addBarValueLabels(ax,x,y,fmt)
%稳定版：逐柱放置数值，正数在上方、负数在下方（防止X/Y维度错误）
if nargin<4,fmt='%.2f';end
axes(ax);hold on;
x=x(:);y=y(:);
n=numel(y);
if numel(x)~=n
    error('addBarValueLabels: length(x)=%d与length(y)=%d不一致。',numel(x),n);
end
yr=max(y)-min(y);
if yr==0
    yr=max(1,abs(max(y)));%退化情形
end
vpad=0.02*yr;%垂直偏移（2%量级）
for i=1:n
    if y(i)>=0
        yPos=y(i)+vpad;
        vAlign='bottom';
    else
        yPos=y(i)-vpad;
        vAlign='top';
    end
    text(x(i),yPos,sprintf(fmt,y(i)),...
        'HorizontalAlignment','center','VerticalAlignment',vAlign);
end
end
function c2=lighten(c,amt)
c2=c+amt*(1-c);
c2=min(max(c2,0),1);
end
function saveFig(f,filename,dpi)
try
    exportgraphics(f,filename,'Resolution',dpi);
catch
    [p,~,~]=fileparts(filename);
    if ~exist(p,'dir'),mkdir(p);end
    print(f,filename,'-dpng',sprintf('-r%d',dpi));
end
end
function writetable_safe(T,xlsx,sheetname)
sheet=regexprep(sheetname,'[:\\/\?\*\[\]]','_');
try
    writetable(T,xlsx,'Sheet',sheet,'WriteMode','overwritesheet');
catch
    writetable(T,xlsx,'Sheet',sheet);
end
end
function [nmLong,nmShort]=strategyNameCN(I1,I2,IF,D)
yn=@(x)ternary(x,'是','否');
nmLong=['预检1=' yn(I1) '，预检2=' yn(I2) '，终检=' yn(IF) '，拆解回用=' yn(D)];
nmShort=['预1-' yn(I1) ' | 预2-' yn(I2) ' | 终-' yn(IF) ' | 拆-' yn(D)];
end
function out=ternary(cond,a,b)
if cond~=0,out=a;else,out=b;end
end
function fname=chooseChineseFont()
candidates={'Microsoft YaHei','SimHei','PingFang SC','Heiti SC','Songti SC',...
    'Noto Sans CJK SC','Source Han Sans SC','Hiragino Sans GB'};
avail=listfonts;
fname=get(groot,'defaultAxesFontName');
for i=1:numel(candidates)
    if any(strcmp(avail,candidates{i}))
        fname=candidates{i};
        return;
    end
end
end
function [idx]=pareto_front(cost,profit)
%目标：成本最小、利润最大非支配点
n=numel(cost);
isPareto=true(n,1);
for i=1:n
    if ~isPareto(i),continue;end
    dom=(cost<=cost(i)&profit>=profit(i))&...
        ((cost<cost(i))|(profit>profit(i)));
    if any(dom),isPareto(i)=false;end
end
idx=find(isPareto);
end
function T=bit_marginal_effect(best_bits,policies_sorted,prof_arr,name_short)
%基于最优策略，逐一翻转I1/I2/IF/D，找到对应策略的利润并给出利润
bitNames={'I1','I2','IF','D'};
best_profit=prof_arr(1);
rows=cell(4,3);%Bit/利润/目标策略名
for b=1:4
    target=best_bits;target(b)=1-target(b);
    match=find(all(policies_sorted==target,2),1,'first');
    if isempty(match)
        dlt=NaN;tname='未找到';
    else
        dlt=prof_arr(match)-best_profit;
        tname=name_short{match};
    end
    rows{b,1}=bitNames{b};
    rows{b,2}=dlt;
    rows{b,3}=tname;
end
T=cell2table(rows,'VariableNames',{'Bit','DeltaProfit','TargetName'});
end
function [vals,names]=topK(x,labels,K)
[~,ord]=sort(x(:),'descend');
K=min(K,numel(ord));
idx=ord(1:K);
vals=x(idx);names=labels(idx);
end