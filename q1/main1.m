%SPRT设计（准则C自动选p1；纯文本绘图；导出XLSX与PNG）
clear;clc;close all;
%全局排版
set(groot,"defaultAxesFontName","Times New Roman",...
           "defaultTextInterpreter","none",...
           "defaultAxesTickLabelInterpreter","none",...
           "defaultLegendInterpreter","none");
%参数区 
p0=0.10;%AQL/原假设次品率
alpha=0.05;%I类错误（H0为真时误拒率）目标：~5%
beta=0.10;%II类错误（p=p1时误收率）目标：~10%
%准则C：业务+成本双约束（用于自动选p1）
delta_min_business=0.05;%最小"明显超标"幅度：从p0+delta起搜索p1
asn_budget_p0=200;%期望样本数预算（在p0下）
asn_budget_p1=200;%期望样本数预算（在p1下）
%p1搜索网格
p1_upper_cap=0.30;
p1_step=1e-3;
%截尾样本量（固定样本兜底阈值会用到；也决定图表计算范围）
Nmax=500;
%灰区处理策略（到达Nmax且未触线）：'accept'或'reject'
gray_policy='accept';
%输出控制
print_full_table=true;%true=全部1..Nmax行；false=节选
table_sample_n=[1 2 3 4 5 10 20 30 50 100 150 200 Nmax];
%保存文件名
xlsname=sprintf('SPRT_数据与检验总结表_Nmax=%d.xlsx',Nmax);
fig1_name=sprintf('图1_决策边界_Nmax=%d.png',Nmax);
fig2_name=sprintf('图2_OC曲线_Nmax=%d.png',Nmax);
fig3_name=sprintf('图3_ASN曲线_Nmax=%d.png',Nmax);
fig4_name=sprintf('图4_p1灵敏度_Nmax=%d.png',Nmax);
%可选仿真验证（关闭则不跑）
do_sim_verify=true;nsim=20000;
%自动选p1（准则C） 
p1_grid=(p0+delta_min_business):p1_step:p1_upper_cap;
assert(~isempty(p1_grid),'p1搜索区间为空，请检查delta_min_business/p1_upper_cap设置。');
Bu=log((1-beta)/alpha);
Bl=log(beta/(1-alpha));
best=struct('p1',NaN,'ASN0',Inf,'ASN1',Inf,'mu0',[],'mu1',[],'a',[],'b',[]);
for p1=p1_grid
    [ASN0,ASN1,mu0,mu1,a,b]=asn_wald_approx(p0,p1,alpha,beta,Bu,Bl);
    if ASN0<=asn_budget_p0&&ASN1<=asn_budget_p1
        if isnan(best.p1)||p1<best.p1||(abs(p1-best.p1)<1e-12&&max(ASN0,ASN1)<max(best.ASN0,best.ASN1))
            best=struct('p1',p1,'ASN0',ASN0,'ASN1',ASN1,'mu0',mu0,'mu1',mu1,'a',a,'b',b);
        end
    end
end
if isnan(best.p1)
    %若预算太紧，选"越贴近预算越好"的p1
    best_cost=Inf;
    for p1=p1_grid
        [ASN0,ASN1,mu0,mu1,a,b]=asn_wald_approx(p0,p1,alpha,beta,Bu,Bl);
        over0=max(0,ASN0-asn_budget_p0);
        over1=max(0,ASN1-asn_budget_p1);
        cost=over0+over1;
        if cost<best_cost||(abs(cost-best_cost)<1e-9&&(isnan(best.p1)||p1<best.p1))
            best_cost=cost;
            best=struct('p1',p1,'ASN0',ASN0,'ASN1',ASN1,'mu0',mu0,'mu1',mu1,'a',a,'b',b);
        end
    end
end
p1=best.p1;a=best.a;b=best.b;
%是否绘制并导出对数似然路径 
make_llr_path=true;%打开/关闭路径图
ptrue_for_path=[p0,p1];%为哪些真实次品率绘制路径（可自行增删）
fig5_name=sprintf('图5_对数似然路径_Nmax%d.png',Nmax);
rng(1);%让示例路径可复现；如要每次随机，注释掉这一行
%生成SPRT阈值表（1..Nmax） 
n_vec=(1:Nmax).';
acc_le=floor((Bl-n_vec.*b)./(a-b));%x <= acc_le -> 接受
rej_ge=ceil((Bu-n_vec.*b)./(a-b));%x >= rej_ge -> 拒绝
acc_le(acc_le< -1)=-1;
rej_ge(rej_ge>n_vec+1)=n_vec(rej_ge>n_vec+1)+1;
%最早可能早停（分开给出）
idx_acc=find(acc_le>=0,1,'first');%首次"可接受"的n
idx_rej=find(rej_ge<=n_vec,1,'first');%首次"可拒绝"的n
%固定样本兜底（n=Nmax）的二项阈值（稳定算法） 
%<= k_acc接收；>= k_rej拒绝（中间灰区按gray_policy处理）
k_acc=binom_quantile_left(Nmax,p0,1-beta);%最小k使P(X <= k) >= 1-β
t95=binom_quantile_left(Nmax,p0,1-alpha);%最小t使P(X <= t) >= 1-α
k_rej=t95+1;%最小k使P(X >= k) <= α
%误差与曲线（OC与ASN，基于动态规划精确计算） 
%用DP精确计算截尾SPRT在给定p下的拒收概率与平均样本数
p_grid=linspace(max(1e-6,p0-0.08),min(0.45,p1+0.12),61).';
OC_rej=zeros(numel(p_grid),1);
ASN_p=zeros(numel(p_grid),1);
for i=1:numel(p_grid)
    [OC_rej(i),ASN_p(i)]=oc_asn_trunc_sprt(p_grid(i),acc_le,rej_ge,Nmax,k_acc,k_rej,gray_policy);
end
%目标点的实际误差与EN
[alpha_hat,EN_p0]=oc_asn_trunc_sprt(p0,acc_le,rej_ge,Nmax,k_acc,k_rej,gray_policy);
[rej_p1,EN_p1]=oc_asn_trunc_sprt(p1,acc_le,rej_ge,Nmax,k_acc,k_rej,gray_policy);
beta_hat=1-rej_p1;
%p1灵敏度分析（误差与ASN随p1的变化） 
p1_sense_grid=(p0+delta_min_business):max(2e-3,p1_step):(p0+0.20);%可按需拉宽
p1_sense_grid=p1_sense_grid(p1_sense_grid<0.5);
sense_out=[];%列：p1, ASN0, ASN1, alpha_hat, beta_hat
for p1c=p1_sense_grid
    a_c=log(p1c/p0);
    b_c=log((1-p1c)/(1-p0));
    acc_c=floor((Bl-n_vec.*b_c)./(a_c-b_c));
    rej_c=ceil((Bu-n_vec.*b_c)./(a_c-b_c));
    acc_c(acc_c< -1)=-1;
    rej_c(rej_c>n_vec+1)=n_vec(rej_c>n_vec+1)+1;
    k_acc_c=binom_quantile_left(Nmax,p0,1-beta);
    t95_c=binom_quantile_left(Nmax,p0,1-alpha);
    k_rej_c=t95_c+1;
    [alpha_hat_c,EN0_c]=oc_asn_trunc_sprt(p0,acc_c,rej_c,Nmax,k_acc_c,k_rej_c,gray_policy);
    [rej_p1_c,EN1_c]=oc_asn_trunc_sprt(p1c,acc_c,rej_c,Nmax,k_acc_c,k_rej_c,gray_policy);
    beta_hat_c=1-rej_p1_c;
    sense_out=[sense_out;[p1c,EN0_c,EN1_c,alpha_hat_c,beta_hat_c]];
end
%打印关键文字结果 
fprintf('\n== SPRT设计（准则C + Nmax=%d）==\n',Nmax);
fprintf('H0: p = %.4f  vs  H1: p = %.4f\n',p0,p1);
fprintf('风险控制: alpha=%.3f (拒收把握约%.1f%%),  beta=%.3f (接收把握约%.1f%%)\n',alpha,100*(1-alpha),beta,100*(1-beta));
fprintf('对数门限: Bu=%.4f,  Bl=%.4f\n',Bu,Bl);
fprintf('单样本对数似然增量: a=%.4f (次品),  b=%.4f (合格)\n',a,b);
fprintf('Wald近似ASN: E[N|p0]=%.1f,  E[N|p1]=%.1f\n',best.ASN0,best.ASN1);
%阈值表
fprintf('\nSPRT阈值表（x <= 接受；x >= 拒绝；其余继续）\n');
Tfull=table(n_vec,acc_le,rej_ge,'VariableNames',{'已检件数','接收若x_le','拒绝若x_ge'});
if print_full_table
    disp(Tfull);
else
    disp(Tfull(ismember(n_vec,unique(min(max(table_sample_n,1),Nmax))),:));
end
%最早可能早停
fprintf('最早可能早停 \n');
if ~isempty(idx_acc)
    fprintf('最早可接受: n = %d 时，若x <= %d即可接受。\n',idx_acc,acc_le(idx_acc));
else
    fprintf('在n <= %d的范围内未出现可接受的阈值。\n',Nmax);
end
if ~isempty(idx_rej)
    fprintf('最早可拒绝: n = %d 时，若x >= %d即可拒绝。\n',idx_rej,rej_ge(idx_rej));
else
    fprintf('在n <= %d的范围内未出现可拒绝的阈值。\n',Nmax);
end
%一句话方案（固定样本兜底）
fprintf('\n一句话方案（固定样本兜底，n=%d）\n',Nmax);
if strcmpi(gray_policy,'accept')
    fprintf('在%.0f%%/%.0f%%信度下，抽%d件：x <= %d接收；x >= %d拒绝；灰区按接收处理。\n',100*(1-beta),100*(1-alpha),Nmax,k_acc,k_rej);
else
    fprintf('在%.0f%%/%.0f%%信度下，抽%d件：x <= %d接收；x >= %d拒绝；灰区按拒绝处理。\n',100*(1-beta),100*(1-alpha),Nmax,k_acc,k_rej);
end
%误差与EN实际值
fprintf('\n实际误差与平均样本数（考虑截尾与灰区策略）\n');
fprintf('实际alpha^ (在p0=%.3f) = %.3f；实际1-beta^ (在p1=%.3f) = %.3f\n',p0,alpha_hat,p1,1-beta_hat);
fprintf('E[N|p0] = %.1f；E[N|p1] = %.1f\n',EN_p0,EN_p1);
%作图并保存PNG（纯文本） 
%8.1决策边界
fig1=figure('Color','w');ax1=axes(fig1);hold(ax1,'on');grid(ax1,'on');
plot(ax1,n_vec,acc_le,'-','LineWidth',1.6);
plot(ax1,n_vec,rej_ge,'-','LineWidth',1.6);
xlabel(ax1,'已检件数n');
ylabel(ax1,'累计次品数x');
title(ax1,{...
  'SPRT决策边界（x<=接收；x>=拒绝）',...
  sprintf('H0: p=%.3f,  H1: p=%.3f,  Nmax=%d',p0,p1,Nmax)...
});
legend(ax1,{'接收边界（低线）','拒绝边界（高线）'},'Location','northwest');
try ax1.Toolbar.Visible='off';catch,end
exportgraphics(fig1,fig1_name,'Resolution',300);
%8.2 OC曲线（拒收概率）
fig2=figure('Color','w');ax2=axes(fig2);hold(ax2,'on');grid(ax2,'on');
plot(ax2,p_grid,OC_rej,'LineWidth',1.6);
yline(ax2,alpha,'--','alpha目标','LabelVerticalAlignment','bottom');
xline(ax2,p0,':','p0','LabelVerticalAlignment','bottom');
xline(ax2,p1,':','p1','LabelVerticalAlignment','bottom');
scatter(ax2,p0,alpha_hat,36,'filled');
scatter(ax2,p1,1-beta_hat,36,'filled');
xlabel(ax2,'真实次品率p');
ylabel(ax2,'拒收概率OC(p)');
title(ax2,{...
  'OC曲线（操作特性）',...
  sprintf('alpha^=%.3f (p0=%.3f),  1-beta^=%.3f (p1=%.3f)',alpha_hat,p0,1-beta_hat,p1)...
});
try ax2.Toolbar.Visible='off';catch,end
exportgraphics(fig2,fig2_name,'Resolution',300);
%8.3 ASN曲线（平均样本数）
fig3=figure('Color','w');ax3=axes(fig3);hold(ax3,'on');grid(ax3,'on');
plot(ax3,p_grid,ASN_p,'LineWidth',1.6);
scatter(ax3,p0,EN_p0,36,'filled');
scatter(ax3,p1,EN_p1,36,'filled');
xlabel(ax3,'真实次品率p');
ylabel(ax3,'平均样本数E[N]');
title(ax3,{...
  'ASN曲线（截尾SPRT）',...
  sprintf('E[N|p0]=%.1f,  E[N|p1]=%.1f',EN_p0,EN_p1)...
});
try ax3.Toolbar.Visible='off';catch,end
exportgraphics(fig3,fig3_name,'Resolution',300);
%8.4 p1灵敏度图
fig4=figure('Color','w');tlo=tiledlayout(fig4,2,1,'TileSpacing','compact','Padding','compact');
%(a)实际alpha^（p0下拒收概率）
ax4a=nexttile(tlo);hold(ax4a,'on');grid(ax4a,'on');
plot(ax4a,sense_out(:,1),sense_out(:,4),'LineWidth',1.6);%alpha_hat
yline(ax4a,alpha,'--','alpha目标');
xlabel(ax4a,'设计备择点p1');
ylabel(ax4a,'实际alpha^');
title(ax4a,'p1灵敏度：实际alpha^(p0)');
%(b)实际beta^（p1下接受概率）
ax4b=nexttile(tlo);hold(ax4b,'on');grid(ax4b,'on');
plot(ax4b,sense_out(:,1),sense_out(:,5),'LineWidth',1.6);%beta_hat
yline(ax4b,beta,'--','beta目标');
xlabel(ax4b,'设计备择点p1');
ylabel(ax4b,'实际beta^');
title(ax4b,'p1灵敏度：实际beta^(p1)');
try
    ax4a.Toolbar.Visible='off';ax4b.Toolbar.Visible='off';
catch
end
%对数似然路径（Sn），并保存PNG
if make_llr_path
    %多路径用多子图展示
    nPaths=numel(ptrue_for_path);
    fig5=figure('Color','w');
    tll=tiledlayout(fig5,nPaths,1,'TileSpacing','compact','Padding','compact');
    for ii=1:nPaths
        ptrue=ptrue_for_path(ii);
        [Sn_vec,x_vec,n_hit,decision]=simulate_llr_path_with_fallback(...
            ptrue,p0,p1,Bu,Bl,Nmax,k_acc,k_rej,gray_policy);
        ax=nexttile(tll);hold(ax,'on');grid(ax,'on');
        plot(ax,1:numel(Sn_vec),Sn_vec,'-','LineWidth',1.6);
        yline(ax,Bu,'--','Bu（上门限）','LabelVerticalAlignment','bottom');
        yline(ax,Bl,':','Bl（下门限）','LabelVerticalAlignment','bottom');
        %标出停止点
        scatter(ax,n_hit,Sn_vec(end),36,'filled');
        xlabel(ax,'检验轮次n');
        ylabel(ax,'累计对数似然Sn');
        title(ax,sprintf('对数似然路径（p=%.3f）：停止于n=%d，结论=%s',ptrue,n_hit,decision));
        legend(ax,{'Sn（逐步累加）','上门限Bu','下门限Bl','停止点'},'Location','best');
        try ax.Toolbar.Visible='off';catch,end
        %写入Excel（每条路径一张表）
        Tllr=table((1:numel(Sn_vec))',x_vec(:),Sn_vec(:),...
                     repmat(Bu,numel(Sn_vec),1),repmat(Bl,numel(Sn_vec),1),...
                     'VariableNames',{'n','x','Sn','Bu','Bl'});
        sheetname=sprintf('LLR路径_p=%.3f',ptrue);
        writetable(Tllr,xlsname,'Sheet',sheetname);
    end
    %图像导出
    exportgraphics(fig5,fig5_name,'Resolution',300);
end

exportgraphics(fig4,fig4_name,'Resolution',300);
%写入Excel（多工作表） 
%表1：阈值表
writetable(Tfull,xlsname,'Sheet','阈值表');
%表2：OC与ASN曲线数据
Tcurve=table(p_grid,OC_rej,ASN_p,'VariableNames',{'p','OC_reject','ASN'});
writetable(Tcurve,xlsname,'Sheet','OC_ASN_曲线');
%表3：p1灵敏度
Tsense=table(sense_out(:,1),sense_out(:,2),sense_out(:,3),sense_out(:,4),sense_out(:,5),...
    'VariableNames',{'p1','E_N_p0','E_N_p1','alpha_hat','beta_hat'});
writetable(Tsense,xlsname,'Sheet','p1_灵敏度');
%表4：汇总
SumNames={'p0','p1','alpha目标','beta目标','Bu','Bl','a','b','k_acc','k_rej',...
             'alpha^实际','beta^实际','E[N|p0]','E[N|p1]','Nmax','灰区策略'};
SumValues={p0,p1,alpha,beta,Bu,Bl,a,b,k_acc,k_rej,alpha_hat,beta_hat,EN_p0,EN_p1,Nmax,gray_policy};
Tsum=table(SumNames',SumValues','VariableNames',{'名称','取值'});
writetable(Tsum,xlsname,'Sheet','汇总');
fprintf('\n数据已写入：%s\n',xlsname);
fprintf('图片已保存：%s, %s, %s, %s\n',fig1_name,fig2_name,fig3_name,fig4_name);
%蒙特卡洛仿真验证误差与EN 
if do_sim_verify
    [rej_rate_p0,Nstop_p0]=simulate_sprt(p0,p0,p1,alpha,beta,Bu,Bl,Nmax,k_acc,k_rej,nsim,gray_policy);
    [rej_rate_p1,Nstop_p1]=simulate_sprt(p1,p0,p1,alpha,beta,Bu,Bl,Nmax,k_acc,k_rej,nsim,gray_policy);
    fprintf('\n蒙特卡洛仿真（nsim=%d）\n',nsim);
    fprintf('p0=%.3f:  拒收率=%.3f (理想=alpha^=%.3f)， E[N]=%.1f\n',p0,rej_rate_p0,alpha_hat,mean(Nstop_p0));
    fprintf('p1=%.3f:  拒收率=%.3f (理想=1-beta^=%.3f)， E[N]=%.1f\n',p1,rej_rate_p1,1-beta_hat,mean(Nstop_p1));
end
%辅助函数
function [ASN0,ASN1,mu0,mu1,a,b]=asn_wald_approx(p0,p1,alpha,beta,Bu,Bl)
    a=log(p1/p0);
    b=log((1-p1)/(1-p0));
    mu0=p0*a+(1-p0)*b;%<0
    mu1=p1*a+(1-p1)*b;%>0
    Eamp0=alpha*Bu+(1-alpha)*(-Bl);
    Eamp1=(1-beta)*Bu+beta*(-Bl);
    ASN0=Eamp0/(-mu0);
    ASN1=Eamp1/(mu1);
end
%稳定计算：二项分布左侧分位点（最小k使P(X <= k) >= q）
function k=binom_quantile_left(n,p,q)
    if q<=0
        k=0;return;
    end
    if q>=1
        k=n;return;
    end
    lo=0;hi=n;
    while lo<hi
        mid=floor((lo+hi)/2);
        Fmid=binom_cdf_via_betainc(mid,n,p);
        if Fmid>=q
            hi=mid;
        else
            lo=mid+1;
        end
    end
    k=lo;
end
function F=binom_cdf_via_betainc(k,n,p)
    if k<0
        F=0.0;return;
    end
    if k>=n
        F=1.0;return;
    end
    %关系：F(k) = I_{1-p}(n-k, k+1)
    F=betainc(1-p,n-k,k+1);
end
%DP精确计算截尾SPRT的拒收概率与平均样本数（给定p）
function [rej_prob,EN]=oc_asn_trunc_sprt(p,acc_le,rej_ge,Nmax,k_acc,k_rej,gray_policy)
    %cont_n(x)表示到第n次检验后、尚未停止、累计次品数为x的概率分布
    cont=1;%n=0时，x=0的概率为1
    rej_prob=0.0;
    EN=0.0;
    for n=1:Nmax
        %由n-1步分布卷积得到n步分布
        cont_next=[cont*(1-p),0]+[0,cont*p];%长度n+1，索引对应x=0..n
        %本步触发接受/拒绝的概率质量
        acc_thr=acc_le(n);%可能是-1（无接收线）
        rej_thr=rej_ge(n);%可能是n+1（无拒绝线）
        mass_acc=0.0;mass_rej=0.0;
        if acc_thr>=0
            mass_acc=sum(cont_next(1:(acc_thr+1)));
            cont_next(1:(acc_thr+1))=0;
        end
        if rej_thr<=n
            mass_rej=sum(cont_next((rej_thr+1):end));
            cont_next((rej_thr+1):end)=0;
        end
        stop_mass=mass_acc+mass_rej;
        rej_prob=rej_prob+mass_rej;
        EN=EN+n*stop_mass;
        cont=cont_next;%更新"继续"区域分布
        %到达Nmax后兜底判定
        if n==Nmax
            if strcmpi(gray_policy,'accept')
                mass_rej_final=sum(cont((k_rej+1):end));
                mass_acc_final=1-rej_prob-(EN/n)-mass_rej_final;%剩余都算接受
            else
                mass_acc_final=sum(cont(1:(k_acc+1)));
                mass_rej_final=1-rej_prob-(EN/n)-mass_acc_final;%剩余都算拒绝
            end
            rej_prob=rej_prob+mass_rej_final;
            EN=EN+n*(mass_rej_final+mass_acc_final);
        end
    end
end
%蒙特卡洛仿真
function [rej_rate,Nstop]=simulate_sprt(p_true,p0,p1,alpha,beta,Bu,Bl,Nmax,k_acc,k_rej,nsim,gray_policy)
    a=log(p1/p0);b=log((1-p1)/(1-p0));
    rej=0;Nstop=zeros(nsim,1);
    for s=1:nsim
        Sn=0.0;x=0;
        for n=1:Nmax
            if rand<p_true
                Sn=Sn+a;x=x+1;
            else
                Sn=Sn+b;
            end
            if Sn>=Bu
                rej=rej+1;Nstop(s)=n;break;
            elseif Sn<=Bl
                Nstop(s)=n;break;
            elseif n==Nmax
                if x>=k_rej
                    rej=rej+1;Nstop(s)=n;break;
                elseif x<=k_acc
                    Nstop(s)=n;break;
                else
                    %灰区按策略
                    if strcmpi(gray_policy,'reject'),rej=rej+1;end
                    Nstop(s)=n;break;
                end
            end
        end
    end
    rej_rate=rej/nsim;
end
function [Sn_vec,x_vec,n_hit,decision]=simulate_llr_path_with_fallback(...
        p_true,p0,p1,Bu,Bl,Nmax,k_acc,k_rej,gray_policy)
    %返回：逐步Sn、逐步次品数x、停止时刻n_hit、决策（ACCEPT/REJECT/GRAY-ACCEPT/GRAY-REJECT）
    a=log(p1/p0);
    b=log((1-p1)/(1-p0));
    Sn=0.0;x=0;
    Sn_vec=zeros(Nmax,1);
    x_vec=zeros(Nmax,1);
    decision='CONTINUE';
    n_hit=Nmax;
    for n=1:Nmax
        if rand<p_true
            Sn=Sn+a;x=x+1;
        else
            Sn=Sn+b;
        end
        Sn_vec(n)=Sn;
        x_vec(n)=x;
        if Sn>=Bu
            n_hit=n;decision='REJECT';
            Sn_vec=Sn_vec(1:n);x_vec=x_vec(1:n);
            return;
        elseif Sn<=Bl
            n_hit=n;decision='ACCEPT';
            Sn_vec=Sn_vec(1:n);x_vec=x_vec(1:n);
            return;
        end
    end
    %若到Nmax仍未触线，则按固定样本兜底
    if x>=k_rej
        decision='REJECT';
    elseif x<=k_acc
        decision='ACCEPT';
    else
        if strcmpi(gray_policy,'accept')
            decision='GRAY-ACCEPT';
        else
            decision='GRAY-REJECT';
        end
    end
end