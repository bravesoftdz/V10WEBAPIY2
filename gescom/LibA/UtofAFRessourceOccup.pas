{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 06/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFRESSOU_OCCUP
Mots clefs ... : TOF:AFOCCUPRESSOU
*****************************************************************}
Unit UtofAFRessourceOccup;

Interface

Uses StdCtrls,Controls,Classes,Windows,sysutils,Graphics,
     UTOF,uTob,stat,HCtrls,HEnt1,UTobView,ParamSoc,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,dbtables,FE_Main,
{$ENDIF}
     DicoAf,UTofAfBaseCodeAffaire,Afplanning,AFPlanningCst,EntGC,UtilRessource;

Type
  TOF_AFOCCUPRESSOU = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      fPeriodicite,fSqlDate,fUnite : String;
      fDtDebut,fDtfin :TDateTime;
      fBRealise : Boolean;
      fTobCritere,fTobAffiche : Tob;
      fColTitre : TStringList;
      fTobViewer :TTobViewer;
      procedure LoadPlanning;
      procedure LoadRessource;
      procedure InitColonne;
      function  FormatColonne(pDtColonne:TdateTime;pStReal:string):string;
      procedure PutColQte  (ptobRess,ptobCol:TOB;pStReal:string);
      procedure AjoutQte   (pStPeriode,pStReal:string ; ptobCol,ptobRess :TOB);
      function  ConvertQte (pDQte:Double) :Double;
      function  DatefinCol (pDateDeb: Tdatetime):Tdatetime;
      procedure TVOnDblClickCell(Sender: TObject);
  end;

  Procedure AFLanceFiche_OccupRessMois;
  Procedure AFLanceFiche_OccupRessSem;

Implementation

const MaxColMois = 12;                //Nombre limité de colonnes Mois
      MaxColSem  = 18;                //Nombre limité de colonnes Semaine
      EntReal : string = ' (R)';      // Ajout sur l'entête des colonnes Réalisé
      LargCol_Ress       : integer = 60;   // Largeur colonne Ressource
      LargCol_Aff        : integer = 110;  // Largeur colonne Affaire
      LargCol_NumTache  : integer = 60;   // Largeur colonne Numtache
      LargCol_LibTache  : integer = 110;  // Largeur colonne Libellé Tache
      LargCol_Cumul      : integer = 70;   // Largeur colonne Cumul
      LargCol_Pla        : integer = 45;   // Largeur colonne quantité planifiée
      LargCol_Rea        : integer = 60;   // Largeur colonne quantité réalisée

      TexteMessage: array[1..1] of string 	= (
          {1}        'Vous n''avez pas de ressources planifiées pour cette période ');

Procedure AFLanceFiche_OccupRessMois;
begin
  AGLLanceFiche ('AFF','AFRESSOU_OCCUP','','','MOIS');
end;

Procedure AFLanceFiche_OccupRessSem;
begin
  AGLLanceFiche ('AFF','AFRESSOU_OCCUP','','','SEM');
end;

procedure TOF_AFOCCUPRESSOU.OnArgument (S : String ) ;
var vDtDeb :TDateTime;
Begin
  Inherited;
  fPeriodicite := Trim(ReadTokenSt(S));
  SetControlVisible('PAVANCE', false);
  SetControlVisible('BParamListe', false);
  SetControlVisible('BPresentation', false);  
  fColTitre := TStringList.Create;
  fTobCritere := TOB.create('Les_Criteres', nil, -1);
  fTobAffiche := TOB.create('Les_Mois', nil, -1);
  TFStat(Ecran).LaTOB :=  fTobAffiche;
  fTobViewer := TTobViewer(getcontrol('TV'));
  fTobViewer.OnDblClick:= TVOnDblClickCell ;
  vDtDeb := strtodate(getcontroltext ('DATEPLANIF'));
  if (fPeriodicite='SEM') then
  begin
    vDtDeb := now;
    while (DayOfWeek (vDtDeb)<>2) do  vDtDeb:= vDtDeb-1;
    setcontroltext ('DATEPLANIF',Datetostr(vDtDeb));
    setcontroltext ('DATEPLANIF_',Datetostr(vDtDeb+41));
  end
  else setcontroltext ('DATEPLANIF_',Datetostr(FinDeMois(PlusMois (vDtDeb,5))));
{$IFDEF EAGLCLIENT}
  setControlVisible('BIMPRIMER', false);
{$ENDIF}
  SetControlText ('UNITETEMPS',VH_GC.AFMesureActivite);
end;

procedure TOF_AFOCCUPRESSOU.OnLoad;
var vDtFin,vDtDebPla,vDtFinPla,vDtDebRea,vDtFinRea:TDatetime;
    vSqlDateRea :string;
begin
  inherited;
  fTobAffiche.cleardetail; fTobCritere.cleardetail; fColTitre.Clear;
  vDtDebPla := strtodate(getcontroltext ('DATEPLANIF'));
  vDtFinPla := strtodate(getcontroltext ('DATEPLANIF_'));
  if (vDtDebPla = iDate1900) then
  begin
     if (fPeriodicite='SEM') then vDtDebPla := now
     else vDtDebPla := DebutDeMois(now);
     setcontroltext ('DATEPLANIF',Datetostr(vDtDebPla));
  end;
  fDtDebut := vDtDebPla; fDtFin := vDtFinPla;
  fBRealise := (Tcheckbox(GetControl('OCCUPREA')).state=cbchecked);
  fUnite := getcontroltext ('UNITETEMPS');
  if fBRealise then
  begin
    vDtDebRea := strtodate(getcontroltext ('DATEREAL'));
    vDtFinRea := strtodate(getcontroltext ('DATEREAL_'));
    if (vDtDebRea <> iDate1900) then
    begin
      if (vDtDebRea < fDtDebut) then fDtDebut := vDtDebRea;
      vSqlDateRea := ' AND APL_DATEDEBREAL >= "'+USDateTime(vDtDebRea)+'"';
    end;
    if (vDtFinRea <> iDate2099) then
    begin
      if (vDtFinRea > fDtFin) then fDtFin := vDtFinRea;
      vSqlDateRea := vSqlDateRea + ' AND APL_DATEFINREAL <= "'+USDateTime(vDtFinRea)+'"';
    end;
  end;
  vDtFin := DatefinCol(fDtDebut);
  if (vDtDebPla <> iDate1900) and (fDtfin > vDtFin) or (fDtDebut > fDtfin) then
  begin
    setcontroltext ('DATEPLANIF_',Datetostr(vDtFin));
    fDtfin := vDtFin; vDtFinPla := vDtFin;
  end;
  fSqlDate := ' AND APL_DATEDEBPLA >= "'+USDateTime(vDtDebPla)+'"'+
              ' AND APL_DATEFINPLA <= "'+USDateTime(vDtFinPla)+'"'+ vSqlDateRea;
end;

procedure TOF_AFOCCUPRESSOU.OnUpdate ;
begin
  Inherited ;
  InitColonne;
  LoadPlanning;
  LoadRessource;
end ;

procedure TOF_AFOCCUPRESSOU.OnDisplay;
var j :integer;
begin
  inherited;
  if fTobCritere.detail.count=0 then exit;
  SetControlVisible('BParamListe', false);
  SetControlVisible('BPresentation', false);
  with fTobViewer do
  begin
    Colwidths[ColIndex('APL_RESSOURCE')]:=0;
    Colwidths[ColIndex('APL_AFFAIRE')]:=LargCol_Aff;
    Colwidths[ColIndex('APL_NUMEROTACHE')]:=0;
    Colwidths[ColIndex('AFF_LIBELLE')]:=LargCol_Aff;
    Colwidths[ColIndex('ARS_LIBELLE')]:=LargCol_Aff;
    Colwidths[ColIndex('ATA_LIBELLETACHE1')]:=LargCol_Aff;
    Colwidths[ColIndex('CumulPla')]:=LargCol_Cumul;
    Colmask [ColIndex('CumulPla')] := '##0.00';
    ColCaption[ColIndex('APL_NUMEROTACHE')]:= 'Tâche';
    ColCaption[ColIndex('AFF_LIBELLE')]:= 'Descriptif de l''affaire';
    ColCaption[ColIndex('ARS_LIBELLE')]:= 'Ressource';
    ColCaption[ColIndex('ATA_LIBELLETACHE1')]:= 'Descriptif de la tâche';
    ColCaption[ColIndex('CumulPla')] := 'Total Planifié';
    if fBrealise then
    begin
      Colwidths[ColIndex('CumulRea')]:=LargCol_Cumul;
      ColFontTitle[ColIndex('CumulRea')].Color :=clBlue;
      ColFontData [ColIndex('CumulRea')].Color :=clBlue;
      ColCaption[ColIndex('CumulRea')] := 'Total Réalisé';
      Colmask [ColIndex('CumulRea')] := '##0.00';
      for j := 0 to fColTitre.Count - 1 do
      begin
        ColCaption[j+8] := fColTitre[j];
        Colmask [j+8] := '##0.00';
        if ((j mod 2)<>0) then
        begin
          ColFontTitle[j+8].Color :=clBlue;
          ColFontData[j+8].Color :=clBlue;
          Colwidths[j+8] := LargCol_Rea;
        end else Colwidths[j+8] := LargCol_Pla;
      end;
    end else
    begin
      for j := 0 to fColTitre.Count - 1 do
      begin
        ColCaption[j+7] := fColTitre[j];
        Colwidths[j+7] := LargCol_Pla;
        colmask [j+7] := '##0.00';
      end;
    end;
  end;
end;

procedure TOF_AFOCCUPRESSOU.OnClose;
begin
  Inherited;
  fTobCritere.cleardetail; fTobCritere.Free; fTobCritere := nil;
  fTobAffiche.cleardetail; fTobAffiche.free; fTobAffiche := nil;
  fColTitre.clear;fColTitre.Free;
end;

procedure TOF_AFOCCUPRESSOU.InitColonne;
var vDtEncours :Tdatetime;
    vInNbColonne :Integer;
begin
  vDtEncours := fDtDebut; vInNbColonne := 0;
  if (fPeriodicite='SEM') then
  begin
    while (DayOfWeek (vDtEncours)<>2) do  vDtEncours:= vDtEncours-1;
    while (vDtEncours <= fDtfin) and (vInNbColonne < MaxColSem) do
    begin
      inc(vInNbColonne);
      fColTitre.Add (FormatColonne(vDtEncours,''));
      if fBRealise then fColTitre.Add (FormatColonne(vDtEncours,EntReal));
      vDtEncours := PlusDate (vDtEncours,1,'S');
    end;
  end else
  begin
    vDtEncours := DebutDeMois (fDtDebut);
    while (vDtEncours <= fDtfin) and (vInNbColonne < MaxColMois) do
    begin
      inc(vInNbColonne);
      fColTitre.Add (FormatColonne(vDtEncours,''));
      if fBRealise then fColTitre.Add (FormatColonne(vDtEncours,EntReal));
      vDtEncours := DebutDeMois(PlusMois (vDtEncours,1));
    end;
  end;
end;

function TOF_AFOCCUPRESSOU.DatefinCol (pDateDeb: Tdatetime):Tdatetime;
begin
  if (fPeriodicite='SEM') then result := PlusDate (pDateDeb,MaxColSem,'S')
  else  result := PlusMois (pDateDeb,MaxColMois);
end;

{***********A.G.L.***********************************************
Auteur  ...... : AB
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Remplir la tob virtuelle avec les données du planning
Mots clefs ... :
*****************************************************************}

procedure TOF_AFOCCUPRESSOU.LoadPlanning;
var vSt,vStCritere : String;
    vQR : TQuery;
begin
  vStCritere := trim(TFStat(Ecran).stsql);
  if (vStCritere = '') then vStCritere := ' WHERE ';
  vSt := 'SELECT APL_RESSOURCE,APL_AFFAIRE,APL_FONCTION,APL_NUMEROTACHE,APL_DATEDEBPLA,APL_DATEFINPLA,APL_QTEPLANIFUREF';
  if (fBrealise) then vSt := vSt + ',APL_DATEDEBREAL,APL_DATEFINREAL,APL_QTEREALUREF ';
  vSt := vSt + ' FROM AFPLANNING,TACHE,RESSOURCE ';
  vSt := vSt + vStCritere + fSqlDate + ' AND APL_RESSOURCE<>"" ';
  vSt := vSt + ' AND APL_AFFAIRE = ATA_AFFAIRE AND APL_NUMEROTACHE = ATA_NUMEROTACHE';
  vSt := vSt + ' AND APL_RESSOURCE=ARS_RESSOURCE';
  vSt := vSt + ' ORDER BY APL_RESSOURCE,APL_AFFAIRE,APL_NUMEROTACHE ';

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then fTobCritere.LoadDetailDB('LIGNEPLANNING','','',vQr,False,True)
    else PGIInfoAf( textemessage[1], Ecran.Caption);
  finally
    if vQr <> nil then Ferme(vQr);
  end;
end;

{***********A.G.L.***********************************************************
Auteur  ...... : AB
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Lignes affichées du planning groupées par ressource,affaire
Mots clefs ... :
******************************************************************************}
procedure TOF_AFOCCUPRESSOU.LoadRessource;
var
  vSt,vStRess,vStAff,vStTache,vStCritere : String;
  vQR : TQuery;
  i,j : Integer;
  vTobLigne,vTC : Tob;
begin
  if fTobCritere.detail.count=0 then exit;
  vStCritere := trim(TFStat(Ecran).stsql);
  if (vStCritere = '') then vStCritere := ' WHERE ';
  vSt := 'SELECT APL_RESSOURCE,ARS_LIBELLE,APL_AFFAIRE,AFF_LIBELLE,APL_NUMEROTACHE,ATA_LIBELLETACHE1,SUM(APL_QTEPLANIFUREF) AS CumulPla';
  if (fBrealise) then  vSt := vSt + ',SUM(APL_QTEREALUREF) AS CumulRea';
  vSt := vSt + ' FROM AFPLANNING,AFFAIRE, TACHE,RESSOURCE ';
  vSt := vSt + vStCritere + fSqlDate + ' AND APL_RESSOURCE<>"" ';
  vSt := vSt + ' AND APL_AFFAIRE = ATA_AFFAIRE AND ATA_NUMEROTACHE = APL_NUMEROTACHE';
  vSt := vSt + ' AND APL_AFFAIRE = AFF_AFFAIRE AND APL_RESSOURCE=ARS_RESSOURCE';
  vSt := vSt + ' GROUP BY APL_RESSOURCE,APL_AFFAIRE,APL_NUMEROTACHE,ARS_LIBELLE,AFF_LIBELLE,ATA_LIBELLETACHE1 ';
  vSt := vSt + ' ORDER BY APL_RESSOURCE,APL_AFFAIRE,APL_NUMEROTACHE ';

  vQr := nil;
  Try
  vQR := OpenSql (vSt,True);
  if Not vQR.Eof then
    begin
      fTobAffiche.LoadDetailDB('AfficheRessource','','',vQr,False,True);
      for j := 0 to fColTitre.Count - 1 do
        fTobAffiche.Detail[0].AddChampSupValeur (fColTitre[j],0.0,True);
      for i := 0 to fTobAffiche.detail.Count - 1 do
      begin
        vTobLigne := fTobAffiche.Detail[i];
        vStRess := vTobLigne.getvalue('APL_RESSOURCE');
        vStAff  := vTobLigne.getvalue('APL_AFFAIRE');
        vStTache := vTobLigne.getvalue('APL_NUMEROTACHE');
        vTobLigne.putvalue('CumulPla', ConvertQte(vTobLigne.getvalue('CumulPla')));
        if fBRealise then
           vTobLigne.putvalue('CumulRea', ConvertQte(vTobLigne.getvalue('CumulRea')));
        vTC := fTobCritere.FindFirst(['APL_RESSOURCE','APL_AFFAIRE','APL_NUMEROTACHE'],[vStRess,vStAff,vStTache],TRUE) ;
        While vTC<>Nil do
        begin
          PutColQte (vTC,vTobLigne,'');
          if fBRealise then PutColQte (vTC,vTobLigne,EntReal);
          vTC:=fTobCritere.FindNext(['APL_RESSOURCE','APL_AFFAIRE','APL_NUMEROTACHE'],[vStRess,vStAff,vStTache],TRUE) ;
        end;
      end;
    end;
  finally
    if vQr <> nil then Ferme(vQr);
  end;
end;

procedure TOF_AFOCCUPRESSOU.PutColQte (ptobRess,ptobCol:TOB;pStReal:string);
var vStPeriodeD,vStPeriodeF : String;
    vDtEnc,vDtFin :TdateTime;
begin
  if (pStReal = '') then
  begin
  vDtEnc := ptobRess.getvalue('APL_DATEDEBPLA');
  vDtFin := ptobRess.getvalue('APL_DATEFINPLA');
  end else
  begin
    vDtEnc := ptobRess.getvalue('APL_DATEDEBREAL');
    vDtFin := ptobRess.getvalue('APL_DATEFINREAL');
    if (vDtEnc = iDate1900) or (vDtFin = iDate1900) or (vDtEnc < fDtDebut) or (vDtEnc > fDtFin) then exit;
  end;
  vStPeriodeD := FormatColonne(vDtEnc,pStReal);
  vStPeriodeF := FormatColonne(vDtFin,pStReal);
  if (vStPeriodeD = vStPeriodeF)  then AjoutQte (vStPeriodeD,pStReal,ptobCol,ptobRess)
  else
  begin
    while (vDtEnc <= vDtFin) do
    begin
      AjoutQte (vStPeriodeD,pStReal,ptobCol,nil);
      vDtEnc := vDtEnc+1;
      vStPeriodeD := FormatColonne(vDtEnc,pStReal);
    end;
//    if (vStPeriodeD = vStPeriodeF)  then AjoutQte (vStPeriodeD,pStReal,ptobCol,nil);
  end;
end;

procedure TOF_AFOCCUPRESSOU.AjoutQte (pStPeriode,pStReal:string ; ptobCol,ptobRess :TOB);
var vQte :double;
begin
  if (fColTitre.IndexOf(pStPeriode) = -1) then exit;
  if (ptobRess = nil) then vQte := ConversionUnite('J',VH_GC.AFMESUREACTIVITE,1)
  else if (pStReal <> '') then vQte := ptobRess.getvalue('APL_QTEREALUREF')
  else vQte := ptobRess.getvalue('APL_QTEPLANIFUREF');
  ptobCol.putvalue( pStPeriode,ptobCol.getvalue (pStPeriode) + ConvertQte(vQte));
end;

function TOF_AFOCCUPRESSOU.ConvertQte(pDQte:Double) :Double;
begin
  Result := pDQte ;
  if (trim(fUnite) <> '') and(fUnite <> VH_GC.AFMESUREACTIVITE) then
    Result := ConversionUnite(VH_GC.AFMESUREACTIVITE,fUnite ,pDQte);
end;

function TOF_AFOCCUPRESSOU.FormatColonne(pDtColonne:TdateTime;pStReal:string):string;
begin
  if (fPeriodicite='SEM') then result := 'Sem '+IntToStr(NumSemaine(pDtColonne)) + pStReal
  else result := formatDateTime('mmmyy',pDtColonne) + pStReal;
end;

procedure TOF_AFOCCUPRESSOU.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Begin
  Aff   := THEdit(GetControl('APL_AFFAIRE'));
  Aff1  := THEdit(GetControl('APL_AFFAIRE1'));
  Aff2  := THEdit(GetControl('APL_AFFAIRE2'));
  Aff3  := THEdit(GetControl('APL_AFFAIRE3'));
  Aff4  := THEdit(GetControl('APL_AVENANT'));
  Tiers := THEdit(GetControl('APL_TIERS'));
End;

procedure TOF_AFOCCUPRESSOU.TVOnDblClickCell(Sender: TObject );
var vStWhere :string;
begin
with fTobViewer do
  begin
    if (ColName[CurrentCol]='AFF_LIBELLE')  then
      V_PGI.DispatchTT( 5,taModif,AsString[ColIndex('APL_AFFAIRE'),CurrentRow], '', 'STATUT:AFF;ETAT:ENC')
    else
    begin  
      vStWhere := ' APL_AFFAIRE = "'+AsString[ColIndex('APL_AFFAIRE'),CurrentRow]+'"';
      ExecPlanning('153202',DateToStr(fDtDebut),vStWhere, '', '-');
    end;                                                      
  end;
end;

Initialization
  registerclasses ([TOF_AFOCCUPRESSOU]) ;
end.



