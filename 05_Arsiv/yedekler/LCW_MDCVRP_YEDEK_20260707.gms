$TITLE LC Waikiki MDCVRP - Nano Model (2 Depo, 8 Magaza, 2 Arac)

* ==============================================================================
* 1. KUMELERIN TANIMLANMASI (SETS)
* ==============================================================================
SET I "Tum noktalar (Depolar ve Magazalar)";
SET T "Arac tipleri" / Panelvan, Kamyon, Tir / ;
SET K "Filodaki araclar" / k1, k2 / ;
SET Map_TK(T,K) "t tipi ile k aracinin eslesmesi" ;

$GDXIN LCW_GAMS_Data.gdx
$LOAD I
$GDXIN

ALIAS (I, j);
ALIAS (I, h);

SET D(I) "Depo noktalari (D kumesi)";
SET M_set(I) "Magaza noktalari (M kumesi)";

D(I) = YES$(ORD(I) <= 2);
M_set(I) = YES$(ORD(I) > 2);

Map_TK('Panelvan', K)$(ORD(K) = 1) = YES;
Map_TK('Kamyon', K)$(ORD(K) = 2) = YES;

* ==============================================================================
* 2. PARAMETRELER
* ==============================================================================
PARAMETER
    DIST(I,j), DUR(I,j), dist_p(I,j), dur_p(I,j), c(I,j,K), r_t(T),
    q(I), Q_cap(K), F_cost(K), W_cost(T), s(I), delta(I,T) ;

$GDXIN LCW_GAMS_Data.gdx
$LOAD DIST, DUR
$GDXIN

dist_p(I,j) = DIST(I,j);
dur_p(I,j) = DUR(I,j) / 60 ;

SCALAR f / 40.0 / ;
SCALAR TLmax / 9.0 / ;

r_t('Panelvan') = 0.10;
r_t('Kamyon')   = 0.22;
r_t('Tir')      = 0.35;

q(I)$M_set(I) = 150;
q(I)$(M_set(I) AND (MOD(ORD(I), 3) = 0)) = 80;
q(I)$(M_set(I) AND (MOD(ORD(I), 3) = 1)) = 150;
q(I)$(M_set(I) AND (MOD(ORD(I), 3) = 2)) = 320;

LOOP(Map_TK(T,K),
    if(SAMEAS(T,'Panelvan'), Q_cap(K) = 1000;  F_cost(K) = 800;);
    if(SAMEAS(T,'Kamyon'),   Q_cap(K) = 2000;  F_cost(K) = 1600;);
);

W_cost('Panelvan') = 3000;
W_cost('Kamyon')   = 6000;
W_cost('Tir')      = 12000;

s(I)$M_set(I) = 0.5;

delta(I, T)$M_set(I) = 1;
delta(I, 'Tir')$(M_set(I) AND (MOD(ORD(I), 4) = 0)) = 0;

LOOP(Map_TK(T,K),
    c(I,j,K) = f * r_t(T) * dist_p(I,j);
);

* ==============================================================================
* 3. MODEL VE DENKLEMLER
* ==============================================================================
VARIABLES Z, u(I,K), n(T) ;
POSITIVE VARIABLES u;
INTEGER VARIABLES n;
BINARY VARIABLES x(I,j,K), y(I,K) ;

EQUATIONS
    OBJ_FUNC, K1_Ziyaret(I), K2_Akis(I,K), K3_Atama(I,K), K4_TekDepo(K),
    K5_Kapasite(K), K6_Zaman(K), K7_MTZ(I,I,K), K8_YukAlt(I,K),
    K8_YukUst(I,K), K9_Fiziksel(I,T), K10_Filo(T) ;

OBJ_FUNC.. Z =E= SUM((K,I,j), c(I,j,K) * x(I,j,K)) + SUM((K,I)$D(I), F_cost(K) * y(I,K)) + SUM(T, W_cost(T) * n(T)) ;
K1_Ziyaret(j)$M_set(j).. SUM((K,I), x(I,j,K)) =E= 1 ;
K2_Akis(h,K).. SUM(I, x(I,h,K)) =E= SUM(j, x(h,j,K)) ;
K3_Atama(I,K)$D(I).. SUM(j$M_set(j), x(I,j,K)) =E= y(I,K) ;
K4_TekDepo(K).. SUM(I$D(I), y(I,K)) =L= 1 ;
K5_Kapasite(K).. SUM(j$M_set(j), q(j) * SUM(I, x(I,j,K))) =L= Q_cap(K) ;
K6_Zaman(K).. SUM((I,j), dur_p(I,j) * x(I,j,K)) + SUM(j$M_set(j), s(j) * SUM(I, x(I,j,K))) =L= TLmax ;
K7_MTZ(I,j,K)$(M_set(I) AND M_set(j) AND (ORD(I) <> ORD(j))).. u(I,K) - u(j,K) + Q_cap(K) * x(I,j,K) =L= Q_cap(K) - q(j) ;
K8_YukAlt(j,K)$M_set(j).. u(j,K) =G= q(j) ;
K8_YukUst(j,K)$M_set(j).. u(j,K) =L= Q_cap(K) ;
K9_Fiziksel(j,T)$M_set(j).. SUM(K$Map_TK(T,K), SUM(I, x(I,j,K))) =L= delta(j,T) ;
K10_Filo(T).. SUM(I$D(I), SUM(K$Map_TK(T,K), y(I,K))) =L= n(T) ;

MODEL LCW_MDCVRP /ALL/ ;
OPTION OPTCR = 0.0 ;   
OPTION RESLIM = 300 ;   
SOLVE LCW_MDCVRP USING MIP MINIMIZING Z ;
DISPLAY Z.L, x.L, y.L;
