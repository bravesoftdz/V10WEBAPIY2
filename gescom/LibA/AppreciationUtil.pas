unit AppreciationUtil;

interface
                                                           
Uses HEnt1, HCtrls, UTOB, Ent1,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
{$IFDEF BTP}
     CalcOleGenericBTP,
{$ENDIF}
     LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls, Classes,
     SysUtils, Dialogs,  UtilPGI, AGLInit, EntGC, affaireutil,HMsgBox,
     paramsoc,factactivite,factutil,saisutil,utilarticle,factcomm,
     activiteutil,utilressource;

const
  OffSetApp : integer = 50000;

Type R_BM = RECORD
		lib1,pres1,four1,frais1: String;
    lib2,pres2,four2,frais2: String;
    bm1,bm2 : boolean;
    sens1,sens2 : integer;
    pr_typ1,pr_ress1,pr_ressaff1 : string;
    pr_prod1,pr_prod2 : boolean;
    pr_p1a,pr_p1b,pr_p1c : string;    // priorité  R:ressource, P : Prod , A : autre
    pr_typ2,pr_ress2,pr_ressaff2 : string;
    pr_p2a,pr_p2b,pr_p2c : string;    // priorité  R:ressource, P : Prod , A : autre
    fo_typ1,fo_ress1,fo_ressaff1 : string;
    fo_prod1,fo_prod2 : boolean;
    fo_p1a,fo_p1b,fo_p1c : string;    // priorité  R:ressource, P : Prod , A : autre
    fo_typ2,fo_ress2,fo_ressaff2 : string;
    fo_p2a,fo_p2b,fo_p2c : string;    // priorité  R:ressource, P : Prod , A : autre
    fr_typ1,fr_ress1,fr_ressaff1 : string;
    fr_prod1,fr_prod2 : boolean;
    fr_p1a,fr_p1b,fr_p1c : string;    // priorité  R:ressource, P : Prod , A : autre
    fr_typ2,fr_ress2,fr_ressaff2 : string;
    fr_p2a,fr_p2b,fr_p2c : string;    // priorité  R:ressource, P : Prod , A : autre
    End;

Type PA_BONI = RECORD
    ind,numapprec : integer;
    aff,cli,typart,art,assist,typecl,ddeb,dfin : string;
    mnt : double;
    End;

// Gestion des boni/mali
procedure RecupParamBoniMali(var BM : R_BM);
procedure GenerationBoniMali( TObEch ,TobAff: TOB ; numdoc : string);
procedure RechProdBoniMali(aff,ddeb,dfin : string;mnt : double;Tobprod : TOB);
Procedure  AppelCreatBoniMali(ind,napp : integer;typgen,aff,cli,numdoc,typart,art,assist,typecl,ddeb,dfin : string;
mnt : double ; BM : R_BM;TobAct ,TobProd,TobAff: TOB; ListeDesCles : TStringList) ;

function  CreationBoniMali(ind,napp : integer;typgen,aff,cli,numdoc,typart,art,assist,typecl,ddeb,dfin : string;
mnt : double ; BM : R_BM;TobAct,Tobprod :TOB; ListeDesCles : TStringList) : boolean;

function EclatBoniMaliProd(ind ,napp: integer;typgen,aff,cli,numdoc,typart,art,assist,typecl,ddeb,dfin : string;
mnt : double;BM : R_BM; TobAct,TobProd: TOB; ListeDesCles : TStringList) : boolean;

Procedure ConvertChampTob(nom : string ; DEV :RDEVISE;tobx : TOB);


implementation

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... :   /  /    
Description .. : Recup du paramètrage des Boni/mali
Mots clefs ... : APPRECIATION;GI
*****************************************************************}
procedure RecupParamBoniMali(var BM : R_BM);
var QQ : Tquery;
		zz,zpri : integer;
    TobDet,TobBM : Tob;
    ztyp : string;
BEGIN
  	TobBM := Tob.Create('les boni/mali',NIL,-1);

	BM.pres1:= '';BM.four1:='';BM.frais1:=''; BM.bm1 := false;
  BM.pres2:= '';BM.four2:='';BM.frais2:=''; BM.bm2 := false;
  BM.pr_p1a := '';BM.pr_p1b := '';BM.pr_p1c := '';
  BM.pr_p2a := '';BM.pr_p2b := '';BM.pr_p2c := '';
  BM.fo_p1a := '';BM.fo_p1b := '';BM.fo_p1c := '';
  BM.fo_p2a := '';BM.fo_p2b := '';BM.fo_p2c := '';
  BM.fr_p1a := '';BM.fr_p1b := '';BM.fr_p1c := '';
  BM.fr_p2a := '';BM.fr_p2b := '';BM.fr_p2c := '';

  if (GetParamSoc('SO_AFAPPAVECBM')= false) then  exit;
  QQ := nil;
  Try;
  // SELECT * : peu de champs, peu d'enregistrements
  QQ := OpenSQL('SELECT * FROM PARBONI order by APB_NUMLIG,APB_NUMBONI',true,-1,'',true) ;
  If Not QQ.EOF then TobBM.LoadDetailDB('PARBONI','','',QQ,True);
  Finally
   Ferme(QQ);
  End;


  for zz:=0  to TobBM.Detail.count-1 do
  Begin                                        
  	Tobdet := TobBM.detail[zz];
    if (Tobdet.GetValue('APB_NUMLIG') = 1) then
    Begin
     	BM.bm1 := true;
      if Tobdet.GetValue('APB_SENS') = '-' then BM.sens1 := (-1) else BM.sens1 := (1);
    	if (BM.lib1 = '') then
      	BM.lib1 := TobDet.GetValue('APB_LIBELLE');

      ztyp := Tobdet.GetValue('APB_TYPEARTICLE');


      if (ztyp = 'PRE') then
      begin
      	BM.pres1 := TobDet.GetValue('APB_ARTICLE');
        BM.pr_typ1 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.pr_ress1 := tobdet.GetValue('APB_RESSOURCE');
        BM.pr_ressaff1 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.pr_prod1 := true else  BM.pr_prod1 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.pr_p1a := 'P';
        if zpri = 2 then BM.pr_p1b := 'P';
        if zpri = 3 then BM.pr_p1c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.pr_p1a := 'R';
        if zpri = 2 then BM.pr_p1b := 'R';
        if zpri = 3 then BM.pr_p1c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.pr_p1a := 'A';
        if zpri = 2 then BM.pr_p1b := 'A';
        if zpri = 3 then BM.pr_p1c := 'A';
      end;

      if (ztyp = 'MAR') then
      begin
      	BM.four1 := TobDet.GetValue('APB_ARTICLE');
        BM.fo_typ1 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.fo_ress1 := tobdet.GetValue('APB_RESSOURCE');
        BM.fo_ressaff1 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.fo_prod1 := true else  BM.fo_prod1 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.fo_p1a := 'P';
        if zpri = 2 then BM.fo_p1b := 'P';
        if zpri = 3 then BM.fo_p1c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.fo_p1a := 'R';
        if zpri = 2 then BM.fo_p1b := 'R';
        if zpri = 3 then BM.fo_p1c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.fo_p1a := 'A';
        if zpri = 2 then BM.fo_p1b := 'A';
        if zpri = 3 then BM.fo_p1c := 'A';

      end;

      if (ztyp = 'FRA') then
      begin
      	BM.frais1 := TobDet.GetValue('APB_ARTICLE');
        BM.fr_typ1 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.fr_ress1 := tobdet.GetValue('APB_RESSOURCE');
        BM.fr_ressaff1 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.fr_prod1 := true else  BM.fr_prod1 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.fr_p1a := 'P';
        if zpri = 2 then BM.fr_p1b := 'P';
        if zpri = 3 then BM.fr_p1c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.fr_p1a := 'R';
        if zpri = 2 then BM.fr_p1b := 'R';
        if zpri = 3 then BM.fr_p1c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.fr_p1a := 'A';
        if zpri = 2 then BM.fr_p1b := 'A';
        if zpri = 3 then BM.fr_p1c := 'A';
      end;


    End;

    if (Tobdet.GetValue('APB_NUMLIG') = 2) then
    Begin
    	BM.bm2 := true;
			if Tobdet.GetValue('APB_SENS') = '-' then BM.sens2 := (-1) else BM.sens2 := (1);
      if (BM.lib2 = '') then
      	BM.lib2 := TobDet.GetValue('APB_LIBELLE');

      ztyp := Tobdet.GetValue('APB_TYPEARTICLE');

      if (ztyp = 'MAR') then
      begin
      	BM.four2 := TobDet.GetValue('APB_ARTICLE');
        BM.fo_typ2 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.fo_ress2 := tobdet.GetValue('APB_RESSOURCE');
        BM.fo_ressaff2 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.fo_prod1 := true else  BM.fo_prod1 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.fo_p2a := 'P';
        if zpri = 2 then BM.fo_p2b := 'P';
        if zpri = 3 then BM.fo_p2c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.fo_p2a := 'R';
        if zpri = 2 then BM.fo_p2b := 'R';
        if zpri = 3 then BM.fo_p2c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.fo_p2a := 'A';
        if zpri = 2 then BM.fo_p2b := 'A';
        if zpri = 3 then BM.fo_p2c := 'A';
      end;

      if (ztyp = 'FRA') then
      begin
      	BM.frais2 := TobDet.GetValue('APB_ARTICLE');
        BM.fr_typ2 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.fr_ress2 := tobdet.GetValue('APB_RESSOURCE');
        BM.fr_ressaff2 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.fr_prod2 := true else  BM.fr_prod2 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.fr_p2a := 'P';
        if zpri = 2 then BM.fr_p2b := 'P';
        if zpri = 3 then BM.fr_p2c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.fr_p2a := 'R';
        if zpri = 2 then BM.fr_p2b := 'R';
        if zpri = 3 then BM.fr_p2c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.fr_p2a := 'A';
        if zpri = 2 then BM.fr_p2b := 'A';
        if zpri = 3 then BM.fr_p2c := 'A';
      end;

    
      if (ztyp = 'PRE') then
      begin
      	BM.pres2 := TobDet.GetValue('APB_ARTICLE');
        BM.pr_typ2 := tobdet.GetValue('APB_TYPEARTICLE');
        BM.pr_ress2 := tobdet.GetValue('APB_RESSOURCE');
        BM.pr_ressaff2 := tobdet.GetValue('APB_RESSAFF');
        if (tobdet.GetValue('APB_PRORATA') = 'X') then
          BM.pr_prod2 := true else  BM.pr_prod2 := false;

        zpri := tobDet.GetValue('APB_PRIPRO');
        if zpri = 1 then BM.pr_p2a := 'P';
        if zpri = 2 then BM.pr_p2b := 'P';
        if zpri = 3 then BM.pr_p2c := 'P';
        zpri := tobDet.GetValue('APB_PRIRESS');
        if zpri = 1 then BM.pr_p2a := 'R';
        if zpri = 2 then BM.pr_p2b := 'R';
        if zpri = 3 then BM.pr_p2c := 'R';
        zpri := tobDet.GetValue('APB_PRIRESSAFF');
        if zpri = 1 then BM.pr_p2a := 'A';
        if zpri = 2 then BM.pr_p2b := 'A';
        if zpri = 3 then BM.pr_p2c := 'A';
      end;

    End;

  End;

 TobBM.free;

END;


{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... :   /  /    
Description .. : Génération des boni/mali dans Activte
Mots clefs ... : APPRECIATION;GI
*****************************************************************}
procedure GenerationBoniMali( TObEch ,TobAff : TOB ; numdoc : string);
Var BM : R_BM;
   // type Prest /frais ou four
    mnt : double;
    dfin,ddeb,aff,cli,typgen : string;
    TobAct,Tobprod : TOB;
    napp : integer;
    ListeDesCles : TStringList;

BEGIN

	RecupParamBoniMali(BM);

  TobAct := TOB.Create('lignes d''activité à créer',Nil,-1);
//  TobCleAct := TOB.Create ('liste cles act' , Nil,-1);
  ListeDesCles := TStringList.Create;
  TobProd := Tob.Create('la production',NIL,-1);

  try
  aff 		:=  TobEch.getValue('AFA_AFFAIRE');
  cli 		:=  TobEch.getValue('AFA_TIERS');
  ddeb 		:= 	DateTOStr(idate1900);   					// GMGMGMGMGM à modifier
  dfin 		:= 	TobEch.GetValue('AFA_DATEECHE');
  napp 		:= 	TobEch.GetValue('AFA_NUMECHE');
  typgen 	:= 	TobEch.GetValue('AFA_GENERAUTO');
  Mnt :=0;

  if (BM.pr_p1a = 'P') or (BM.pr_p1b = 'P') or (BM.pr_p2a = 'P') or (BM.pr_p2b = 'P') or
 		 (BM.fr_p1a = 'P') or (BM.fr_p1b = 'P') or (BM.fr_p2a = 'P') or (BM.fr_p2b = 'P') or
     (BM.fo_p1a = 'P') or (BM.fo_p1b = 'P') or (BM.fo_p2a = 'P') or (BM.fo_p2b = 'P') then
  begin
		RechProdBoniMali(aff,ddeb,dfin,mnt,TObProd);
  end;

  mnt := TobEch.getValue('AFA_BONIMALI');
  if (mnt <> 0) then
//    AppelCreatBoniMali(1,napp,typgen,aff,cli,numdoc,'PRE','','','',ddeb,dfin,mnt,BM,TobAct, TobcleAct,tobprod,TobAff);
    AppelCreatBoniMali (1, napp, typgen, aff, cli, numdoc, 'PRE', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);



  mnt := TobEch.getValue('AFA_BM1FO');
  if (mnt <> 0) then
    AppelCreatBoniMali (1, napp, typgen, aff, cli, numdoc, 'MAR', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);

  mnt := TobEch.getValue('AFA_BM1FR');
  if (mnt <> 0) then
    AppelCreatBoniMali (1, napp, typgen, aff, cli, numdoc, 'FRA', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);

  mnt := TobEch.getValue('AFA_BM2PR');
  if (mnt <> 0) then
    AppelCreatBoniMali (2, napp, typgen, aff, cli, numdoc, 'PRE', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);

  mnt := TobEch.getValue('AFA_BM2FO');
  if (mnt <> 0) then
    AppelCreatBoniMali (2, napp, typgen, aff, cli, numdoc, 'MAR', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);

  mnt := TobEch.getValue('AFA_BM2FR');
  if (mnt <> 0) then
    AppelCreatBoniMali (2, napp, typgen, aff, cli, numdoc, 'FRA', '', '', '', ddeb, dfin, mnt, BM, TobAct, tobprod, TobAff, ListeDesCles);


	  // mcd 25/11/02 ne plus utiliserif TobAct.Detail.count > 0 then TobAct.InsertDBTable(Nil);
	if TobAct.Detail.count > 0 then
    TobAct.InsertDB (Nil);

  finally
	TobAct.Free;
  //TobCleAct.Free;
  ListeDesCles.Free;
  TobProd.free;
  end;

END;

Procedure  AppelCreatBoniMali(ind,napp : integer;typgen,aff,cli,numdoc,typart,art,assist,typecl,ddeb,dfin : string;
mnt : double ; BM : R_BM;TobAct, Tobprod ,TobAff: TOB; ListeDesCles : TStringList);
var pa,pb,pc,ress,ressaff,zassist,ztypecl,zart : string;
		ret : boolean;
    wi : integer;
BEGIN

  	if (ind = 1) then
    begin
      if (typart  = 'PRE') then
      begin
      	zart := BM.pres1;
        pa := BM.pr_p1a; pb := BM.pr_p1b ;pc :=BM.pr_p1c;
    		ress := BM.pr_ress1;
      	ressaff := BM.pr_ressaff1;
      end;

      if (typart  = 'FRA') then
      begin
      	zart := BM.frais1;
        pa := BM.fr_p1a; pb := BM.fr_p1b ;pc :=BM.fr_p1c;
    		ress := BM.fr_ress1;
      	ressaff := BM.fr_ressaff1;
      end;

      if (typart  = 'MAR') then
      begin
      	zart := BM.four1;
        pa := BM.fo_p1a; pb := BM.fo_p1b ;pc :=BM.fo_p1c;
    		ress := BM.fo_ress1;
      	ressaff := BM.fo_ressaff1;
      end;
    end
    else
    begin
    	if (typart  = 'PRE') then
      begin
      	zart := BM.pres2;
        pa := BM.pr_p2a; pb := BM.pr_p2b ;pc :=BM.pr_p2c;
        ress := BM.pr_ress2;
        ressaff := BM.pr_ressaff2;
      end;
      if (typart  = 'FRA') then
      begin
      	zart := BM.frais2;
        pa := BM.fr_p2a; pb := BM.fr_p2b ;pc :=BM.fr_p2c;
    		ress := BM.fr_ress2;
      	ressaff := BM.fr_ressaff2;
      end;
      if (typart  = 'MAR') then
      begin
      	zart := BM.four2;
        pa := BM.fo_p2a; pb := BM.fo_p2b ;pc :=BM.fo_p2c;
    		ress := BM.fo_ress2;
      	ressaff := BM.fo_ressaff2;
      end;
    end;


    wi := 1; ret := false;
    while  (wi <= 3) and (not(ret)) do        // boucle sur les 3 priorité
      Begin
      if (wi = 1) then ztypecl := pa;     // priorité 1
      if (wi = 2) then ztypecl := pb;     // priorité 2
      if (wi = 3) then ztypecl := pc;     // priorité 3
      zassist := '';

      if (ztypecl = 'R') or (ztypecl = 'A') then
      begin
        if (ztypecl = 'R') then
          zassist := ress
        else
          zassist := TobAff.GetValue(ressaff);  //recherche responsable affaire
          
        if (zassist <> '') then
          ret := CreationBoniMali (ind, napp, typgen, aff, cli, numdoc, typart, zart, zassist, ztypecl, ddeb, dfin, mnt, BM,
                                    tobact, tobprod, ListeDesCles)
        else
          ret := false ;
      end;

      if (ztypecl = 'P') then
      begin
        ret := EclatBoniMaliProd (ind, napp, typgen, aff, cli, numdoc, typart, zart, zassist, ztypecl, ddeb, dfin, mnt, BM,
                                  tobact, tobprod, ListeDesCles);
      end;

      inc(wi);
    End;

    if not(ret) then  // aucun eclatement n'a abouti
    begin

    end;

    //si aucun eclatement exempl PA = P  (pas de prod)
    //  et responsable affaore = ''
    // que fait -on ?

END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... : 14/08/2001
Description .. : Création des Boni/mali en fonction du paramètrage souhaité
Mots clefs ... : APPRECIATION;GI
*****************************************************************}
function CreationBoniMali (ind, napp : integer; typgen, aff, cli, numdoc, typart, art,assist, typecl, ddeb, dfin : string;
                          mnt : double; BM : R_BM; TobAct, TobProd : TOB; ListeDesCles : TStringList) : boolean;
  var
  typeactivite, folio, part0, part1, part2, part3, avnt, dim : string;
	numligne, sens : integer;
  TobDetAct : TOB;
  DEV : RDEVISE;
begin
	//result := false;
  TobDetAct := TOB.Create ('ACTIVITE', TobAct, -1);

  if ind = 1 then  sens := BM.sens1 else sens := BM.sens2;

	TypeActivite := 'BON';
  Folio := '01';

	TobDetAct.InitValeurs;

  TobDetAct.PutValue ('ACT_TYPEACTIVITE', 'BON');
  TobDetAct.PutValue ('ACT_TYPERESSOURCE', 'SAL');   //tobech
  TobDetAct.PutValue ('ACT_RESSOURCE', assist);
  TobDetAct.PutValue ('ACT_FOLIO', 1);

  TobDetAct.PutValue ('ACT_DATEACTIVITE', strtodate (dfin));
  TobDetAct.PutValue ('ACT_PERIODE', GetPeriode (strtodate (dfin)));
  TobDetAct.PutValue ('ACT_SEMAINE', NumSemaine (strtodate (dfin)));

  TobDetAct.PutValue ('ACT_TYPEARTICLE', typart);

  TobDetAct.PutValue ('ACT_ARTICLE', art);
  dim := '';
  TobDetAct.PutValue ('ACT_CODEARTICLE', trim (CodeArticleGenerique (art, dim, dim, dim, dim, dim)));        //GM
  TobDetAct.PutValue ('ACT_LIBELLE', 'recup boni/mali');


  TobDetAct.PutValue ('ACT_DEVISE', V_PGI.DevisePivot);

  TobDetAct.PutValue ('ACT_TIERS', cli);
  TobDetAct.PutValue ('ACT_AFFAIRE', aff);
  {$IFDEF BTP}
  BTPCodeAffaireDecoupe (Aff, part0, part1, part2, part3, avnt, tamodif, false);
  {$ELSE}
  CodeAffaireDecoupe (Aff, part0, part1, part2, part3, avnt, tamodif, false);
  {$ENDIF}

  TobDetAct.PutValue ('ACT_AFFAIRE0', Part0);
  TobDetAct.PutValue ('ACT_AFFAIRE1', Part1);
  TobDetAct.PutValue ('ACT_AFFAIRE2', Part2);
  TobDetAct.PutValue ('ACT_AFFAIRE3', Part3);
  TobDetAct.PutValue ('ACT_AVENANT', avnt);

	TobDetAct.PutValue ('ACT_UNITE', '');

  TobDetAct.PutValue ('ACT_QTE', 1 * sens);
 //TobDetAct.PutFixedStValue ('ACT_UNITEFAC', StTot, 146,1, tcChaine, false);
  TobDetAct.PutValue ('ACT_QTEFAC', 1 * sens);
  if (TobDetAct.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
    TobDetAct.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TobDetAct.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobDetAct.GetValue ('ACT_QTE')))
  else
    TobDetAct.PutValue ('ACT_QTEUNITEREF', '0');

 	TobDetAct.PutValue ('ACT_DEVISE', V_PGI.DevisePivot);

  // Gestion des prix de revient. On met la même valeur de partout
 // TobDetAct.PutFixedStValue ('ACT_PUPR', StTot, 154,12, tcDouble100, false);
 // TobDetAct.PutFixedStValue ('ACT_PUPRCHARGE', StTot, 154,12, tcDouble100, false);
 // TobDetAct.PutFixedStValue ('ACT_PUPRCHINDIRECT', StTot, 154,12, tcDouble100, false);
 // TobDetAct.PutFixedStValue ('ACT_TOTPR', StTot, 178,12, tcDouble100, false);
  //TobDetAct.PutFixedStValue ('ACT_TOTPRCHARGE', StTot, 178,12, tcDouble100, false);
//  TobDetAct.PutFixedStValue ('ACT_TOTPRCHINDI', StTot, 178,12, tcDouble100, false);

	DEV.Code := V_PGI.DevisePivot;
	GetInfosDevise (DEV);    // attention ne pas appeler à chaque fois
  TobDetAct.PutValue ('ACT_PUVENTE', mnt);
  ConvertChampTob ('ACT_PUVENTE', DEV, TobDetAct);

  TobDetAct.PutValue ('ACT_TOTVENTE', mnt * sens);
  ConvertChampTob ('ACT_TOTVENTE', DEV, TobDetAct);



  TobDetAct.PutValue ('ACT_DATECREATION', Nowh);
  TobDetAct.PutValue ('ACT_DATEMODIF', Nowh);
  //TobDetAct.PutValue ('ACT_REALISABLE','X');

  // VIS  si solde , si partiel , ne doit pas être visé
  if (typgen = 'SOL') then
    begin
      TobDetAct.PutValue ('ACT_ETATVISA', 'VIS');
      TobDetAct.PutValue ('ACT_VISEUR', V_PGI.User);
	    TobDetAct.PutValue ('ACT_DATEVISA', NowH);
      TobDetAct.PutValue ('ACT_ETATVISAFAC', 'VIS');
	    TobDetAct.PutValue ('ACT_VISEURFAC', V_PGI.User);
	    TobDetAct.PutValue ('ACT_DATEVISAFAC', NowH);
      TobDetAct.PutValue ('ACT_ACTIVITEREPRIS', 'N');
    end
  else
    begin
  	  TobDetAct.PutValue ('ACT_ETATVISAFAC', 'ATT');
	    TobDetAct.PutValue ('ACT_VISEURFAC', V_PGI.User);
	    TobDetAct.PutValue ('ACT_DATEVISAFAC', idate1900);

  	  TobDetAct.PutValue ('ACT_ETATVISA', 'VIS');
	    TobDetAct.PutValue ('ACT_VISEUR', V_PGI.User);
	    TobDetAct.PutValue ('ACT_DATEVISA', NowH);
      TobDetAct.PutValue ('ACT_ACTIVITEREPRIS', 'F');
    end;

  TobDetAct.PutValue ('ACT_NUMAPPREC', napp);

  // Num de pièce associée
  TobDetAct.PutValue ('ACT_NUMPIECE', numdoc);

  ////////////// PL le 14/04/03 : modifs changement de clé dans la table ACTIVITE
  // Recherche du num de ligne de la clé 1
  (*  NumLigne := GetNumLigneInCleAct (TobDetAct, TobCleAct);
  if NumLigne = 0 then
    begin
      NumLigne := MaxNumLigneActivite (TobDetAct); // Attention une requête est faite à chaque fois ...
      InsertCleInCleAct (TobDetAct, TobCleAct, NumLigne + 1);
    end;
  TobDetAct.PutValue ('ACT_NUMLIGNE', NumLigne + 1);
  *)

  NumLigne := ProchainIndiceAffaires (TobDetAct.GetValue ('ACT_TYPEACTIVITE'), TobDetAct.GetValue ('ACT_AFFAIRE'), ListeDesCles);
  TobDetAct.PutValue ('ACT_NUMLIGNEUNIQUE', NumLigne);
  TobDetAct.PutValue ('ACT_NUMLIGNE', NumLigne); // La valeur du numéro d'ordre a peu d'importance
  ////////////////// Fin PL le 14/04/03


  result := true; //// a compléter

end;


{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 14/08/2001
Modifié le ... : 20/08/2001
Description .. : Recherche de la prod en vue de la génération des
Suite ........ : Boni/mali
Suite ........ : Eclatement seulement par rapport aux perestations
Mots clefs ... : APPRECIATION;GI
*****************************************************************}
procedure RechProdBoniMali(aff,ddeb,dfin : string;mnt : double;Tobprod : TOB) ;
var QQ :TQuery;
		Req1,Req,zwhere,zwhere2,zorder,{ztyp,}wheredate : String;

BEGIN

// PL le 06/03/02 : INDEX 1       AAAAAAAAAAAAAAAAAAAAAAAAAAAA
  Req1 	 := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE,ACT_RESSOURCE,ACT_TYPEARTICLE FROM ACTIVITE ';
  zwhere := ' WHERE  ACT_TYPEACTIVITE="REA" AND ACT_AFFAIRE="'+aff+  '"';
  wheredate := ' AND ACT_DATEACTIVITE <= "'+ usdatetime(StrTodate(dfin))+'"';
  wheredate := wheredate + ' AND ACT_DATEACTIVITE >= "'+ usdatetime(StrTodate(ddeb))+'"';
  zwhere2 :=  ' AND ACT_TYPEARTICLE="PRE"';
  if (GetParamSoc ('SO_AFVISAACTIVITE') = True) then // PL le 19/09/03 : on ne prend que les lignes visées dans le cas ou le visa sur l'activite est geré dans les paramètres
    zwhere2 :=  zwhere2+ ' AND ACT_ETATVISA="VIS"';

  zorder := '  GROUP BY ACT_RESSOURCE,ACT_TYPEARTICLE';

  Req 	 := Req1+zwhere+wheredate+zwhere2+zorder ;
  QQ := nil;
  Try;
  QQ := OpenSQL(Req,true,-1,'',true) ;
  If Not QQ.EOF then TobProd.LoadDetailDB('ACTIVITE','','',QQ,True);
  Finally
   Ferme(QQ);
  End;

END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 20/08/2001
Modifié le ... :   /  /
Description .. : Generation des Boni mali par rapport à la production
Mots clefs ... : GI;APPRECIATION
*****************************************************************}
Function EclatBoniMaliProd (ind, napp : integer; typgen, aff, cli, numdoc, typart, art, assist, typecl, ddeb, dfin : string;
                            mnt : double; BM : R_BM; TobAct, TobProd : TOB; ListeDesCles : TStringList) : boolean;
var   {vte,}cumtot,coef,mntcalc,cumcalc : double;
    	{wj,}wi: integer;
    	TobDet : TOB;
      //ret : boolean;
//const
//    ttyp: array[1..3] of string  = ('PRE','MAR','FRA');
//    Maxtyp : integer = 3;
// pour l'instant je  récupère le prorata par ressource et par prestation
// et j'applique ce coef aussi sur les frais et fournitures
// voir s'il faut faire un distingo + tard

BEGIN
	result := true;
  cumcalc := 0; //mntcalc := 0;
  CumTot := TOBProd.Somme('ACT_TOTVENTE',['ACT_TYPEARTICLE'],['PRE'],False);

  if cumtot = 0 then        // pas de prod sur la periode, eclatement impossible
  begin
  	result := false;
    exit;
  end;

  // On ne doit pas éclater si la ressource n'a pas de prestation en production

  for wi := 0 to TObProd.detail.count-1 do
  begin
    TobDet := TobProd.detail[wi];

    coef := arrondi ((TobDet.GetValue('ACT_TOTVENTE') / cumtot),2);
//    coef := (TObDet.GetValue('ACT_TOTVENTE') / cumtot);    PL le 31/01/02 à revoir
    assist := TobDet.GetValue('ACT_RESSOURCE');

    if (wi <> TobProd.detail.count - 1) then
      begin
        mntcalc := arrondi ((mnt * coef) , 2);
        cumcalc := cumcalc + mntcalc;
      end
    else
      begin   // derniere ressource 
        mntcalc := mnt - cumcalc;
      end;

    CreationBoniMali (ind, napp, typgen, aff, cli, numdoc, typart, art, assist, typecl, ddeb, dfin, mntcalc, BM, tobact, tobprod, ListeDesCles);

  End;


END;

Procedure ConvertChampTob(nom : string ; DEV :RDEVISE ;tobx : TOB);
var MtDev,MtPivot : double;
BEGIN
  // PL le 24/01/02
  MtPivot:= Tobx.GetValue(nom);
  // C.B 18/06/2003 Suppression contrevaleur
  ConvertPivotToDev(DEV, MtPivot,  MtDev);
  Tobx.PutValue(nom+'DEV',MtDev);
  // Fin PL le 24/01/02
END;

end.
