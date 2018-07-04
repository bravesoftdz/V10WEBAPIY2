unit UtilSynVte;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox, UTOB,
{$IFDEF EAGLCLIENT}
      utileAGL,
{$ELSE}
      db,dbTables,Fiche,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}      
{$ENDIF}
      AglInit,EntGC, HQry,HStatus ;



procedure RempliTOBSynthese(TobTemp : TOB ; var TobSyn : TOB;typsyn :string;Prixttc:boolean; Prixach:boolean; codetrait:string; regroup:string; tri:string; ddeb:TDateTime; dfin:TDateTime; critere : string='');
procedure RempliTousLesJours(TobTemp : TOB ; var TobSyn : TOB); overload;
procedure RempliTousLesJours(TobTemp : TOB ;Annee : Word ; Mois : Word ; var TobSyn : TOB); overload;
procedure RempliTousLesMois(TobTemp : TOB ;Annee : Word ; var TobSyn : TOB);
procedure RecalcStockMois(dep:string; art:string ; familleNiv1:string; familleNiv2:string; familleNiv3:string; dateJ:TDateTime; var Qsto:double; var Qatt:TOB; trait : string);
implementation

procedure RempliTOBSynthese(TobTemp : TOB ; var TobSyn : TOB;typsyn :string;Prixttc:boolean; Prixach:boolean; codetrait:string; regroup:string; tri:string; ddeb:TDateTime; dfin:TDateTime; critere : string='');
var TobPal, TOBT : TOB;
    Annee, Mois, Jour, Semaine, annee_sv, mois_sv : Word;
    i_ind, i : integer;
    qtefac,reelht,reeltc, brutttc, brutht, qteret, mtrettc, mtretht, mtach : double;
    LaDate : TDateTime ;
    rupt1, rupt2, librupt1, librupt2, wclef : string ;
begin
InitMove(TobTemp.Detail.count,'');
for i_ind:=0 to TobTemp.Detail.count-1 do
  begin
  TobT := TobTemp.Detail[i_ind];
  if (codetrait = 'SYN') then
    Ladate := TobT.GetValue('GL_DATEPIECE')
  else begin
    if (TobT.GetValue('GL_DATEPIECE') >= ddeb) and (TobT.GetValue('GL_DATEPIECE') <= dfin) then
      Ladate := TobT.GetValue('GL_DATEPIECE')
    else
      Ladate := ddeb;
  end;
  DecodeDate(LaDate, Annee, Mois, Jour);
  if (typsyn = 'AD') or (typsyn = 'AA') then
    begin
    jour := 1;
    LaDate := Encodedate (Annee, Mois, Jour)
    end;

  Semaine := NumSemaine (LaDate) ;

  librupt2 := '';
  rupt2 := '';
  librupt1 := '';
  rupt1 := '';
  if (codetrait = 'SYN') then
     begin
     if critere = 'AA' then
       TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_RUPT0','GZS_ANNEE','GZS_MOIS'],
       [V_PGI.USer,codetrait,TobT.GetValue('GL_ETABLISSEMENT'),IntToStr(Integer(Annee)),IntToStr(Integer(mois))], False)
     else
       TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_RUPT0','GZS_ANNEE','GZS_MOIS','GZS_JOUR'],
       [V_PGI.USer,codetrait,TobT.GetValue('GL_ETABLISSEMENT'),IntToStr(Integer(Annee)),IntToStr(Integer(mois)),IntToStr(Integer(jour))], False)
     end
  else begin
    if regroup = 'ART' then
       begin
       rupt2 := TOBT.GetValue('GA_CODEARTICLE');
       librupt2 := TOBT.GetValue('GA_CODEARTICLE');
       end
    else if regroup = 'FOU' then
       begin
       rupt2 := TOBT.GetValue('GA_FOURNPRINC');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('T_LIBELLE');
       end
    else if regroup = 'COL' then
       begin
       rupt2 := TOBT.GetValue('GA_COLLECTION');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('COLLEC');
       end
    else if regroup = 'NV1' then
       begin
       rupt2 := TOBT.GetValue('GA_FAMILLENIV1');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('LIBFAMNIV1');
       end
    else if regroup = 'NV2' then
       begin
       rupt2 := TOBT.GetValue('GA_FAMILLENIV2');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('LIBFAMNIV2');
       end
    else if regroup = 'NV3' then
       begin
       rupt2 := TOBT.GetValue('GA_FAMILLENIV3');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('LIBFAMNIV3');
       end
    else if regroup = 'ETA' then
       begin
       rupt2 := TOBT.GetValue('GQ_DEPOT');
       if (rupt2 <> '') then  librupt2 := TOBT.GetValue('ET_LIBELLE');
       end;

    if tri = 'FOU' then
       begin
       rupt1 := TOBT.GetValue('GA_FOURNPRINC');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('T_LIBELLE');
       end
    else if tri = 'COL' then
       begin
       rupt1 := TOBT.GetValue('GA_COLLECTION');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('COLLEC');
       end
    else if tri = 'NV1' then
       begin
       rupt1 := TOBT.GetValue('GA_FAMILLENIV1');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('LIBFAMNIV1');
       end
    else if tri = 'NV2' then
       begin
       rupt1 := TOBT.GetValue('GA_FAMILLENIV2');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('LIBFAMNIV2');
       end
    else if tri = 'NV3' then
       begin
       rupt1 := TOBT.GetValue('GA_FAMILLENIV3');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('LIBFAMNIV3');
       end
    else if tri = 'ETA' then
       begin
       rupt1 := TOBT.GetValue('GQ_DEPOT');
       if (rupt1 <> '') then  librupt1 := TOBT.GetValue('ET_LIBELLE');
       end;
    wclef := TobT.GetValue('GQ_DEPOT') + TobT.GetValue('GA_CODEARTICLE');
    if (codetrait = 'ETV') then  // etat des ventes
      TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_RUPT0','GZS_ANNEE','GZS_MOIS','GZS_SEM','GZS_RUPT1','GZS_RUPT2'],
      [V_PGI.USer,codetrait,wclef,IntToStr(Integer(Annee)),IntToStr(Integer(mois)),IntToStr(Integer(semaine)),rupt1,rupt2], False);
    if (codetrait = 'GRV') then  // Graphe
      TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_ANNEE','GZS_MOIS','GZS_SEM','GZS_RUPT1','GZS_RUPT2'],
      [V_PGI.USer,codetrait,IntToStr(Integer(Annee)),IntToStr(Integer(mois)),IntToStr(Integer(semaine)),rupt1,rupt2], False);
    end;
  MoveCur(False);
  if TobPal = nil then
     begin
     if codetrait = 'SYN' then   // synthese des ventes
       begin
       TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
       TobPal.InitValeurs;
       TobPal.PutValue('GZS_TRAIT',codetrait);
       TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
       TobPal.PutValue('GZS_RUPT0',TOBT.GetValue('GL_ETABLISSEMENT'));
       TobPal.PutValue('GZS_ANNEE',annee);
       TobPal.PutValue('GZS_MOIS',mois);
       TobPal.PutValue('GZS_JOUR',jour);
       TobPal.PutValue('GZS_DATE',Ladate);
       end;
     if (codetrait = 'ETV') or (codetrait = 'GRV') then  // etat des ventes ou Graphe
       begin
       LaDate := ddeb;
       annee_sv := annee;
       mois_sv := mois;
       for i := 1 to 13 do
       begin
         if  Ladate <= dfin then
           begin
           DecodeDate(Ladate, annee, mois, Jour);
           Semaine := NumSemaine (Ladate) ;
           TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
           TobPal.InitValeurs;
           TobPal.PutValue('GZS_TRAIT',codetrait);
           TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
           if (codetrait = 'ETV') then  // Etat
//             TobPal.PutValue('GZS_RUPT0',TOBT.GetValue('GQ_DEPOT'));
             TobPal.PutValue('GZS_RUPT0',wclef);
           if (codetrait = 'GRV') then  // Graphe
             TobPal.PutValue('GZS_RUPT0','');
           TobPal.PutValue('GZS_ANNEE',annee);
           TobPal.PutValue('GZS_MOIS',mois);
           TobPal.PutValue('GZS_SEM',semaine);
           TobPal.PutValue('GZS_JOUR',i);
           TobPal.PutValue('GZS_RUPT1',rupt1);
           TobPal.PutValue('GZS_RUPT2',rupt2);
           TobPal.PutValue('GZS_LIBRUPT1',librupt1);
           TobPal.PutValue('GZS_LIBRUPT2',librupt2);
           TobPal.PutValue('GZS_DATE',Ladate);
           TobPal.PutValue('GZS_QTESTOCK',TOBT.GetValue('STOCK'));
           TobPal.PutValue('GZS_QTEATT',TOBT.GetValue('ATTENDU'));
           end;
         LaDate := PlusMois(ddeb,i) ;
         end;
       if (codetrait = 'ETV') then  // Etat
         TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_RUPT0','GZS_ANNEE','GZS_MOIS','GZS_RUPT1','GZS_RUPT2'],
         [V_PGI.USer,codetrait,wclef,IntToStr(Integer(Annee_sv)),IntToStr(Integer(mois_sv)),rupt1,rupt2], False);
       if (codetrait = 'GRV') then  // Graphe
         TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_TRAIT','GZS_ANNEE','GZS_MOIS','GZS_RUPT1','GZS_RUPT2'],
         [V_PGI.USer,codetrait,IntToStr(Integer(Annee_sv)),IntToStr(Integer(mois_sv)),rupt1,rupt2], False);
       end;
     end;
  qtefac := 0;
  if (TOBT.GetValue('GL_QTEFACT') <> null) then
    qtefac := TOBT.GetValue('GL_QTEFACT');

  if (codetrait = 'ETV') or  (codetrait = 'GRV') then  // Etat ou Graphe
    begin
    TobPal.PutValue('GZS_QTESTOCK',TOBT.GetValue('STOCK'));
    TobPal.PutValue('GZS_QTEATT',TOBT.GetValue('ATTENDU'));
    end;

  if (codetrait = 'SYN') then
    begin
    brutttc := TOBT.GetValue('BRUTTC');
    brutht := TOBT.GetValue('BRUTHT');
    reelht := TOBT.GetValue('GL_TOTALHT');
    reeltc := TOBT.GetValue('GL_TOTALTTC');
    qteret := 0.00;
    mtrettc := 0.00;
    mtretht := 0.00;
    mtach := (TOBT.GetValue('GL_PMAP') * qtefac);
    if (qtefac < 0.00) then
      begin
      qteret := qtefac * (-1);
      mtrettc := reeltc *(-1);
      mtretht := reelht *(-1);
      qtefac := 0.00;
      brutttc := 0.00;
      brutht := 0.00;
      reeltc := 0.00;
      reelht := 0.00;
      end;
    end;
  TobPal.PutValue('GZS_QTEVTE',TobPal.GetValue('GZS_QTEVTE')+qtefac);
  if (codetrait = 'SYN') then
    begin
    TobPal.PutValue('GZS_QTERET',TobPal.GetValue('GZS_QTERET')+ qteret);
    TobPal.PutValue('GZS_QTETOT',TobPal.GetValue('GZS_QTETOT')+ qtefac - qteret);
    TobPal.PutValue('GZS_MTREELHT',TobPal.GetValue('GZS_MTREELHT')+ reelht);
    TobPal.PutValue('GZS_MTREELTC',TobPal.GetValue('GZS_MTREELTC')+ reeltc);
    TobPal.PutValue('GZS_MTRETTC',TobPal.GetValue('GZS_MTRETTC')+mtrettc);
    TobPal.PutValue('GZS_MTRETHT',TobPal.GetValue('GZS_MTRETHT')+ mtretht);

    if Prixttc = True then
      begin
      TobPal.PutValue('GZS_MTBRUT',TobPal.GetValue('GZS_MTBRUT')+brutttc);
      TobPal.PutValue('GZS_MTREM',TobPal.GetValue('GZS_MTREM')+ brutttc - reeltc);
      TobPal.PutValue('GZS_MTREEL',TobPal.GetValue('GZS_MTREEL')+ reeltc);
      TobPal.PutValue('GZS_MTRET',TobPal.GetValue('GZS_MTRET')+mtrettc);
      TobPal.PutValue('GZS_CATOT',TobPal.GetValue('GZS_CATOT')+ reeltc - mtrettc);
      end else
      begin
      TobPal.PutValue('GZS_MTBRUT',TobPal.GetValue('GZS_MTBRUT')+brutht);
      TobPal.PutValue('GZS_MTREM',TobPal.GetValue('GZS_MTREM')+ brutht - reelht);
      TobPal.PutValue('GZS_MTREEL',TobPal.GetValue('GZS_MTREEL')+ reelht) ;
      TobPal.PutValue('GZS_MTRET',TobPal.GetValue('GZS_MTRET')+ mtretht);
      TobPal.PutValue('GZS_CATOT',TobPal.GetValue('GZS_CATOT')+ reelht - mtretht);
      end;

    if Prixach = True then
      begin
      if qteret = 0.00 then
        TobPal.PutValue('GZS_PRIXACH',TobPal.GetValue('GZS_PRIXACH')+ mtach);
      TobPal.PutValue('GZS_PRIXACHTOT',TobPal.GetValue('GZS_PRIXACHTOT')+ mtach);
      end;
    end;
  end;
FiniMove;
end;

procedure RempliTousLesJours(TobTemp : TOB ; var TobSyn : TOB);
var TobPal, TOBT : TOB;
    Annee, Mois, Jour : Word;
    i_ind,y_ind, nbjour : integer;
    LaDate : TDateTime ;
begin
InitMove(TobTemp.Detail.count,'');
for i_ind:=0 to TobTemp.Detail.count-1 do
  begin
  TobT := TobTemp.Detail[i_ind];
  Ladate := TobT.GetValue('GL_DATEPIECE');
  DecodeDate(LaDate, Annee, Mois, Jour);

  TobPal := TobSyn.FindFirst(['GZS_UTILISATEUR','GZS_RUPT0','GZS_ANNEE','GZS_MOIS'],
  [V_PGI.USer,TobT.GetValue('GL_ETABLISSEMENT'),IntToStr(Integer(Annee)),IntToStr(Integer(mois))], False);
  MoveCur(False);
  if TobPal = nil then
    begin
    nbjour := 30;
    if (Mois = 1) or (Mois = 3) or (Mois = 5)
    or (Mois = 7) or (Mois = 8) or (Mois = 10)
    or (Mois = 12) then nbjour := 31;
    if (Mois = 2) then
      begin
      if (isLeapYear(Annee) = True) then nbjour := 29
      else nbjour := 28;
      end;
    for y_ind :=1 to nbjour do
      begin
      TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
      TobPal.InitValeurs;
      TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
      TobPal.PutValue('GZS_TRAIT','SYN');
      TobPal.PutValue('GZS_RUPT0',TOBT.GetValue('GL_ETABLISSEMENT'));
      TobPal.PutValue('GZS_ANNEE',annee);
      TobPal.PutValue('GZS_MOIS',mois);
      TobPal.PutValue('GZS_JOUR',y_ind);
      jour := y_ind;
      LaDate := Encodedate (Annee, Mois, Jour);
      TobPal.PutValue('GZS_DATE',Ladate);
      end;
    end;
  end;
FiniMove;
end;

procedure RempliTousLesJours(TobTemp : TOB ; Annee : Word ; Mois : Word ; var TobSyn : TOB);
var TobPal, TobT : TOB;
    Jour : Word;
    Etab_encours : string;
    i_ind,y_ind, nbjour : integer;
    LaDate : TDateTime ;
begin
InitMove(TobTemp.Detail.count,'');
Etab_encours := '';
nbjour := 30;
if (Mois = 1) or (Mois = 3) or (Mois = 5)
or (Mois = 7) or (Mois = 8) or (Mois = 10)
or (Mois = 12) then nbjour := 31;
if (Mois = 2) then
   begin
   if (isLeapYear(Annee) = True) then nbjour := 29
   else nbjour := 28;
   end;
for i_ind:=0 to TobTemp.Detail.count-1 do
   begin
   MoveCur(False);
   TobT := TobTemp.Detail[i_ind];
   if Etab_encours <> TobT.GetValue('GL_ETABLISSEMENT') then
      begin
      Etab_encours := TobT.GetValue('GL_ETABLISSEMENT');
      for y_ind :=1 to nbjour do
         begin
         TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
         TobPal.InitValeurs;
         TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
         TobPal.PutValue('GZS_TRAIT','SYN');
         TobPal.PutValue('GZS_RUPT0',Etab_encours);
         TobPal.PutValue('GZS_ANNEE',annee);
         TobPal.PutValue('GZS_MOIS',mois);
         TobPal.PutValue('GZS_JOUR',y_ind);
         jour := y_ind;
         LaDate := Encodedate (Annee, Mois, Jour);
         TobPal.PutValue('GZS_DATE',Ladate);
         end;
      for y_ind :=1 to nbjour do
         begin
         TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
         TobPal.InitValeurs;
         TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
         TobPal.PutValue('GZS_TRAIT','SYN');
         TobPal.PutValue('GZS_RUPT0',Etab_encours);
         TobPal.PutValue('GZS_ANNEE',Annee-1);
         TobPal.PutValue('GZS_MOIS',mois);
         TobPal.PutValue('GZS_JOUR',y_ind);
         jour := y_ind;
         LaDate := Encodedate (Annee-1, Mois, Jour);
         TobPal.PutValue('GZS_DATE',Ladate);
         end;
      end;
   end;
FiniMove;
end;

procedure RempliTousLesMois(TobTemp : TOB ;Annee : Word ; var TobSyn : TOB);
var
   TobPal, TobT : TOB;
   Etab_encours : string;
   i_ind, y_ind, nbMois : integer;
   Mois : Word;
   LaDate : TDateTime;
begin
InitMove(TobTemp.Detail.count,'');
Etab_encours := '';
nbMois :=12;
for i_ind:=0 to TobTemp.Detail.count-1 do
   begin
   MoveCur(False);
   TobT := TobTemp.Detail[i_ind];
   if Etab_encours <> TobT.GetValue('GL_ETABLISSEMENT') then
      begin
      Etab_encours := TobT.GetValue('GL_ETABLISSEMENT');
      for y_ind :=1 to nbMois do
         begin
         TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
         TobPal.InitValeurs;
         TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
         TobPal.PutValue('GZS_TRAIT','SYN');
         TobPal.PutValue('GZS_RUPT0',Etab_encours);
         TobPal.PutValue('GZS_ANNEE',annee);
         TobPal.PutValue('GZS_MOIS',y_ind);
         TobPal.PutValue('GZS_JOUR',1);
         Mois := y_ind;
         LaDate := Encodedate (Annee-1, Mois, 1);
         TobPal.PutValue('GZS_DATE',Ladate);
         end;
      for y_ind :=1 to nbMois do
         begin
         TobPal := TOB.Create('GCTMPSYNVTE',TobSyn,-1);
         TobPal.InitValeurs;
         TobPal.PutValue('GZS_UTILISATEUR',V_PGI.USer);
         TobPal.PutValue('GZS_TRAIT','SYN');
         TobPal.PutValue('GZS_RUPT0',Etab_encours);
         TobPal.PutValue('GZS_ANNEE',annee-1);
         TobPal.PutValue('GZS_MOIS',y_ind);
         TobPal.PutValue('GZS_JOUR',1);
         Mois := y_ind;
         LaDate := Encodedate (Annee-1, Mois, 1);
         TobPal.PutValue('GZS_DATE',Ladate);
         end;
      end;
   end;
FiniMove;
end;

procedure RecalcStockMois(dep:string; art:string ; familleNiv1:string; familleNiv2:string; familleNiv3:string; dateJ:TDateTime; var Qsto:double; var Qatt:TOB; trait : string);
var qte_stock : double;
    QLigne : TQuery;
    date_cloture : TDateTime ;
    requeteDateCloture, requeteStock, requeteQteAttendu : string;
    TobLigne,TobL, qte_attendu : TOB;
begin
qte_stock := 0;
qte_attendu := nil;
if trait <> 'ATT' then
   begin
   requeteDateCloture:= 'SELECT MAX(GQ_DATECLOTURE) as DATECLO FROM DISPOART_MODE '+
   'WHERE GA_TYPEARTICLE = "MAR" AND GQ_CLOTURE = "X" AND GQ_DEPOT = "'+dep+'" AND ';
   requeteStock:= 'SELECT GQ_PHYSIQUE as STOCK FROM DISPOART_MODE '+
      'WHERE GA_TYPEARTICLE = "MAR" AND GQ_CLOTURE = "X" AND GQ_DEPOT = "'+dep+'" AND ';
   if art <> '' then
      begin
      requeteDateCloture:=requeteDateCloture + 'GA_CODEARTICLE = "'+art+ '" AND ';
      requeteStock:=requeteStock + 'GA_CODEARTICLE = "'+art + '" AND ';
      end;
   if familleNiv1 <> '' then
      begin
      requeteDateCloture:=requeteDateCloture + 'GA_FAMILLENIV1 = "' + familleNiv1+ '" AND ';
      requeteStock:=requeteStock + 'GA_FAMILLENIV1 = "' + familleNiv1 + '" AND ';
      end;
   if familleNiv2 <> '' then
      begin
      requeteDateCloture:=requeteDateCloture + 'GA_FAMILLENIV2 = "' + familleNiv2 + '" AND ';
      requeteStock:=requeteStock + 'GA_FAMILLENIV2 = "' + familleNiv2 + '" AND ';
      end;
   if familleNiv3 <> '' then
      begin
      requeteDateCloture:=requeteDateCloture + 'GA_FAMILLENIV3 = "' + familleNiv3 + '" AND ';
      requeteStock:=requeteStock + 'GA_FAMILLENIV3 = "' + familleNiv3 + '" AND ';
      end;
   requeteDateCloture:=requeteDateCloture + 'GQ_DATECLOTURE <= "'+USDateTime(dateJ)+'"';
   requeteStock:=requeteStock + 'GQ_DATECLOTURE = "'+USDateTime(date_cloture)+'"';
   QLigne := OpenSQL(requeteDateCloture,True);
   date_cloture := 0;
   if not QLigne.Eof then
      begin
      date_cloture := QLigne.Findfield('DATECLO').AsDateTime;
      Ferme(QLigne);
      QLigne := OpenSQL(requeteStock,True);
      if not QLigne.Eof then
         begin
         qte_stock := QLigne.Findfield('STOCK').AsFloat;
         end;
      end;
   Ferme(QLigne);
   end;

if trait <> 'STO' then
   begin
   QLigne := OpenSQL('SELECT GL_CODEARTICLE, GL_ETABLISSEMENT, '+
   'sum (GL_QTEFACT) as QTE FROM LIGNE '+
   'WHERE GL_NATUREPIECEG = "CF" '+
   'AND GL_CODEARTICLE in ('+art+') AND GL_ETABLISSEMENT in ('+dep+')'+
   ' AND GL_VIVANTE = "X" AND GL_DATEPIECE <= "'+USDateTime(dateJ)+
   '" GROUP BY GL_CODEARTICLE, GL_ETABLISSEMENT', True);
   if not QLigne.Eof then
      begin
      qte_attendu:=TOB.Create('',nil,-1);
      qte_attendu.LoadDetailDB('','','',QLigne,false);
      end;
   Ferme(QLigne);
   end;
Qatt := qte_attendu;
Qsto := qte_stock;
end;

end.
