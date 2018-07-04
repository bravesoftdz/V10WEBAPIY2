{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 03/03/2004
Modifié le ... :   /  /                                   
Description .. : Source TOF de la FICHE : PGEVOLUTIONMASSE ()
Mots clefs ... : TOF;PGEVOLUTIONMASSE
*****************************************************************}
{
 PT1 : 10/10/2005 MF V610 FQ 12614 :
                            1-le traitement est effectué sur la période
                            choisie et non plus sur l'exercice complet.
                            2-Le calcul des effectifs prend en compte soit tous
                            les salariés présents au cours du mois soit uniquement
                            ceux présents à fin de période (Boîte à cocher)
 PT3 : 25/09/2009 SJ FQ n°12376 Suppression du LanceEtatTob
}
Unit PGEVOLUTIONMASSE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     EdtREtat,
     QRS1,//PT3
{$else}
     eQRS1,//PT3
     UtileAGL,
{$ENDIF}
     sysutils,
     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     HQry,   // PT1-2
     ParamDat,
     ParamSoc,
     P5Def,
     UTob,
     PgOutils2,
     UTOF ;

Type
  TOF_PGEVOLUTIONMASSE = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
//PT1    ExDateDeb, ExDateFin, Categ, Annee  : string;
    ExDateDeb, ExDateFin, Categ         : string;
    TobMass,TOBEtat,T                   : TOB;
    ChampRuptOrig                       : array[1..9] of string;
    ChampRupt                           : array[1..9] of string;
    ValeurRupt                          : array[1..9] of string;
    Tablette                            : array[1..9] of string;
    EffPresent                          : array[1..12] of integer;
    NbRupt                              : Integer;
    
    procedure BEtatClick(Sender : TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ControlPeriode(Sender: TObject);
    procedure ChangeExercice(Sender: TObject);
    procedure PgIFValidPeriode;
    procedure ConstruireTob ();

  end ;

Implementation

uses EntPaie;

procedure TOF_PGEVOLUTIONMASSE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONMASSE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONMASSE.OnUpdate ;
begin
  Inherited ;
 BEtatClick(nil);//PT3
end ;

procedure TOF_PGEVOLUTIONMASSE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONMASSE.OnArgument (S : String ) ;
var
  THDateDeb,THDateFin           : THEdit;
  Exercice,Combo                : THValComboBox ;     
  QPeriode,QExer                : TQuery;
  {BEtat                         : TToolBarButton97;}//PT3 mise en commentaire 
  num,i                         : Integer;

begin
  Inherited ;
  SetControlText('DOSSIER',GetParamSocSecur ('SO_LIBELLE',''));
  SetControlChecked('PAIEDECALEE',VH_Paie.PGDecalage);

  QExer:=OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT),MAX(PEX_DATEFIN)'+
                 ',MAX(PEX_ANNEEREFER)'+
                 ' FROM EXERSOCIAL '+
                 'WHERE PEX_ACTIF="X"',true);
  if not QExer.eof then
  begin
    SetControlText('EDEXERSOC',QExer.Fields[0].asstring);
    ExDateDeb :=QExer.Fields[1].asstring;
    ExDateFin :=QExer.Fields[2].asstring;
//PT1    Annee := QExer.Fields[3].asstring;
  end
  else
  begin
    SetControlText('EDEXERSOC','');
    ExDateDeb:=DateToStr(idate1900);
    ExDateFin:=DateToStr(idate1900);
//PT1 Annee := '';
  end;
  Ferme(QExer);

  // d PT1-1
//  THDateDeb:= ThEdit(getcontrol('XX_VARIABLEDEB'));
//  THDateFin:= ThEdit(getcontrol('XX_VARIABLEFIN'));
  THDateDeb:= ThEdit(getcontrol('PHC_DATEDEBUT'));
  THDateFin:= ThEdit(getcontrol('PHC_DATEFIN'));
// f PT1-1

  QPeriode:=OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) '+
                    'FROM PAIEENCOURS '+
                    'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND '+
                    'PPU_DATEFIN<=PEX_DATEFIN '+
                    'WHERE PEX_EXERCICE="'+GetControltext('EDEXERSOC')+'"',TRUE);
  if not QPeriode.eof then
  begin
    ExDateDeb:=DateToStr(QPeriode.Fields[0].AsDateTime);
    ExDateFin:=DateToStr(QPeriode.Fields[1].AsDateTime);
  end;
  Ferme(QPeriode);

  Exercice:= THValComboBox(getcontrol('EDEXERSOC'));
  if (THDateDeb <> nil) and (THDateFin <> nil) and (Exercice<>nil) then
  begin
   THDateDeb.text := ExDateDeb;
   THDateDeb.OnElipsisClick := DateElipsisclick;
   THDateDeb.OnExit := ControlPeriode;
   THDateFin.text := ExDateFin;
   THDateFin.OnElipsisClick := DateElipsisclick;
   THDateFin.OnExit := ControlPeriode;
   Exercice.OnChange := ChangeExercice;
  end;

{  BEtat  := TToolBarButton97(GetControl('BVALIDER'));
  if Betat <> Nil then BEtat.OnClick  := BEtatClick;}//PT3 mise en commentaire 

// affichage dynamique des critères de rupture
  for num := 1 to 9 do
  begin
    VisibiliteChampRupt(IntToStr(num) ,
                        GetControl('VRUPT'+IntToStr(num)),
                        GetControl('TRUPT'+IntToStr(num)));
  end;

  for num := 1 to (VH_Paie.pgNbreStatOrg + VH_Paie.PgNbCombo + 1) do
  begin
    Combo := THValComboBox(GetControl('VRUPT'+IntToStr(num)));
    if Combo <> NIL then
    begin
      Combo.Items.Add('<<Aucun>>');
      Combo.Values.Add('');
      Combo.Items.Add('Etablissement');
      Combo.Values.Add('PHC_ETABLISSEMENT');
      if (VH_Paie.PgNbreStatOrg <> 0) then
      begin
        for i := 1 to VH_Paie.PgNbreStatOrg do
        begin
          case i of
          1 : begin
                Combo.Items.Add(VH_Paie.PgLibelleOrgStat1);
                Combo.Values.Add('PSA_TRAVAILN1');
              end;
          2 : begin
                Combo.Items.Add(VH_Paie.PgLibelleOrgStat2);
                Combo.Values.Add('PSA_TRAVAILN2');
              end;
          3 : begin
                Combo.Items.Add(VH_Paie.PgLibelleOrgStat3);
                Combo.Values.Add('PSA_TRAVAILN3');
              end;
          4 : begin
                Combo.Items.Add(VH_Paie.PgLibelleOrgStat4);
                Combo.Values.Add('PSA_TRAVAILN4');
              end;
          end;
        end;
      end;
      if (VH_Paie.PgNbCombo <> 0 ) then
      begin
        for i := 1 to VH_Paie.PgNbCombo do
        begin
          case i of
          1 : begin
                Combo.Items.Add(VH_Paie.PgLibCombo1);
                Combo.Values.Add('PSA_LIBREPCMB1');
              end;
          2 : begin
                Combo.Items.Add(VH_Paie.PgLibCombo2);
                Combo.Values.Add('PSA_LIBREPCMB2');
              end;
          3 : begin
                Combo.Items.Add(VH_Paie.PgLibCombo3);
                Combo.Values.Add('PSA_LIBREPCMB3');
              end;
          4 : begin
                Combo.Items.Add(VH_Paie.PgLibCombo4);
                Combo.Values.Add('PSA_LIBREPCMB4');
              end;
          end;
        end;
      end;
    end;
  end;
end ;

procedure TOF_PGEVOLUTIONMASSE.OnClose ;
begin
  Inherited ;
 If TobEtat <> Nil Then FreeAndNil(TobEtat);//PT3
end ;

procedure TOF_PGEVOLUTIONMASSE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGEVOLUTIONMASSE.OnCancel () ;
begin
  Inherited ;
end ;
procedure TOF_PGEVOLUTIONMASSE.BEtatClick(Sender : TObject);
var
  Pages                                         : TPageControl;
  num, num2,i                                   : Integer;
  Where,GroupBy,SQL,Wselect,WhereEff            : String;
  QMass,QEff                                    : TQuery;
  TobMassCat                                    : TOB;
  AnneeDebut,Moisdebut,AnneeFin,MoisFin,jj      : word;
  FinMois, aa, mm                               : word;
  StPages,ess:string;
  Tess : TDateTime;
  FinPeriode                    : TCheckBox;                 // PT1-2

begin

If TobEtat <> Nil Then FreeAndNil(TobEtat);//PT3
 // TobEtat := NIL; PT3 mise en commentaire
  Pages := TPageControl(GetControl('Pages'));
// d PT1-1
  Where := '';
  Where := RecupWhereCritere(Pages);
// f  PT1-1

  // Récupération des critères de rupture
  NbRupt := 0;
  for num := 9 downto 1 do
  begin
    if (num <> 1) and (GetControlText('VRUPT'+IntToStr(num)) <> '') then
    begin
      for num2 := num-1 downto 1 do
      // vérif. les niveaux de rupture précédents doivent être renseignés
      begin
        if (GetControlText('VRUPT'+IntToStr(num2)) = '') then
        begin
          PGIBox('Le niveau de rupture '+IntToStr(num2)+' doit être renseigné',Ecran.caption);
          exit;
        end;
      end;
    end;
    if (GetControlText('VRUPT'+IntToStr(num))= '') then continue;

    setControlText('XX_RUPTURE'+IntToStr(num),GetControlText('VRUPT'+IntToStr(num)));
    NbRupt := NbRupt + 1;
  end;
  if (NbRupt > 1) then
  begin
    for num := 1 to NbRupt do
    // Vérif. si pls fois la même rupture
    begin
      if (GetControlText('VRUPT'+IntToStr(num)) <> '') then
      begin
        for num2 := 1 to NbRupt do
        begin
          if (GetControlText('VRUPT'+IntToStr(num)) = GetControlText('VRUPT'+IntToStr(num2))) and(num <> num2) then
          begin
            PGIBox('Un critère de rupture a été choisi plusieurs fois. Taitement impossible',Ecran.caption);
            exit;
          end;
        end;
      end;
    end;
  end;

  // construction de la requête
// d PT1
// d PT1-2
  Where := '';
  Where := RecupWhereCritere(Pages);
  if (Where <> '') then
     Where := Where + ' AND PHC_CUMULPAIE="01" AND PSA_DADSPROF <> "" '
  else
//    Where := 'WHERE PHC_CUMULPAIE="01" AND YEAR(PHC_DATEFIN) = "'+Annee+'" '+
    Where := 'WHERE PHC_CUMULPAIE="01" AND PSA_DADSPROF <> ""';
// f PT1

  WhereEff := ' AND PSA_DADSPROF <> ""';
  GroupBy := 'GROUP BY ';
  Wselect := '';

  for num := 1 to 9 do
  // initialisations des variables de travail permettant la gestion des ruptures
  begin
    ChampRupt[num] := '';
    ValeurRupt[num] := '';
    Tablette[num] := '';
  end;

  if (NbRupt = 0) then
  begin
    SetControlText('VRUPT1','établissement');
    Where := Where + ' AND PHC_ETABLISSEMENT <> ""';
    WhereEff := WhereEff + ' AND PSA_ETABLISSEMENT <> ""';
    GroupBy := GroupBy + 'PHC_ETABLISSEMENT';
    Wselect := Wselect + ',PHC_ETABLISSEMENT';

    ChampRuptOrig[1] := 'PHC_ETABLISSEMENT';
    ChampRupt[1] := 'RUPTURE1';
    for num := 2 to 9 do
      ChampRupt[num] := 'RUPTURE'+IntToStr(num);
  end
  else
  begin
    for num := 1 to NbRupt do
    begin
      Where := Where + ' AND '+ GetControlText('XX_RUPTURE'+IntToStr(num)) + ' <> ""';
      if (GetControlText('XX_RUPTURE'+IntToStr(num)) = 'PHC_ETABLISSEMENT') then
        WhereEff := WhereEff + ' AND PSA_ETABLISSEMENT <> ""'
      else
        WhereEff := WhereEff + ' AND '+ GetControlText('XX_RUPTURE'+IntToStr(num)) + ' <> ""';
      Wselect := Wselect + ','+ GetControlText('XX_RUPTURE'+IntToStr(num));
      ChampRuptOrig[num] := GetControlText('XX_RUPTURE'+IntToStr(num));
      ChampRupt[num] := 'RUPTURE'+IntToStr(num);
      if (ChampRuptOrig[num] <> 'PHC_ETABLISSEMENT') then
        Tablette[num] := 'PG' + Copy(ChampRuptOrig[num],5,length(ChampRuptOrig[num]));
    end;
    if (NbRupt <> 9) then
    begin
      for num := NbRupt+1 to 9 do
        ChampRupt[num] := 'RUPTURE'+IntToStr(num);
    end;
    for num := 1 to NbRupt do
    begin
      if (num <> 1) then
      begin
        GroupBy := GroupBy + ',';
      end;
      GroupBy := GroupBy + GetControlText('XX_RUPTURE'+IntToStr(num));
    end;
  end;
  GroupBy := GroupBy +',MOnTH(PHC_DATEFIN),PSA_DADSPROF,ET_LIBELLE';
  SQL := 'SELECT sum(PHC_MONTANT) MT, max(PHC_MONTANT) MAXMT, '+
         'min (PHC_MONTANT) MINMT,PSA_DADSPROF CATEGORIE';
  SQL := SQL + ',Count (Distinct (PHC_SALARIE)) EFF';   //@@effectif
  SQL := SQL + Wselect;
  SQL := SQL + ',ET_LIBELLE, MOnTH(PHC_DATEFIN) MOIS FROM HISTOCUMSAL '+
               'lEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PHC_ETABLISSEMENT '+
               'LEFT JOIN SALARIES ON PSA_SALARIE=PHC_SALARIE ';
  SQL := SQL + Where + ' '+GroupBy;
  QMass := OpenSql(SQL,TRUE);
  TobMassCat := Tob.Create('Les Masses salariales',Nil,-1);
  TobMassCat.LoadDetailDB('Les Masses salariales','','',QMass,False);
  Ferme(QMass);

  // calcul de l'effectif présent pour la période
  // salariés pour lesquels les critères de rupture sont renseignés
// d PT1-1
//  Decodedate(StrToDate(ExDateDeb),AnneeDebut,MoisDebut,jj);
//  Decodedate(StrToDate(ExDateFin),AnneeFin,MoisFin,jj);
  DecodeDate(StrToDate(GetControlText('PHC_DATEDEBUT')),AnneeDebut,MoisDebut,jj);
  DecodeDate(StrToDate(GetControlText('PHC_DATEFIN')),AnneeFin,MoisFin,jj);

//  if (VH_Paie.PGDecalage = TRUE) then MoisFin := 12;
// f PT1-1

  FinPeriode := TCheckBox(GetControl('FINPERIODE')); //PT1-2

  i := MoisDebut;
  while i <= MoisFin do
  begin
    ess := '01/'+ColleZeroDevant(i,2)+'/'+IntToStr(AnneeDebut);
    Tess := StrToDate(ess);

    Tess := FinDeMois(Tess);
    ess := DateToStr(Tess);
    Decodedate(Tess,aa, mm,FinMois);
    SQL := '';
{
    SQL := 'SELECT COUNT(PSA_SALARIE) FROM SALARIES '+
           'WHERE PSA_DATEENTREE <= "'+ColleZeroDevant(i,2)+'/'+
           IntToStr(FinMois)+'/'+IntToStr(AnneeDebut)+'" AND '+
           '(PSA_DATESORTIE IS NULL OR PSA_DATESORTIE = "01/01/1900" OR '+
           'PSA_DATESORTIE >= "'+ColleZeroDevant(i,2)+'/01/'+IntToStr(AnneeDebut)+'") AND '+
           'PSA_SUSPENSIONPAIE <> "X"'+WhereEff;}
    SQL := 'SELECT COUNT(PSA_SALARIE) FROM SALARIES ';
    SQL :=  SQL + 'WHERE PSA_DATEENTREE <= "'+ColleZeroDevant(i,2)+'/'+
                   IntToStr(FinMois)+'/'+IntToStr(AnneeDebut)+'" AND ';
      if (FinPeriode <> nil) then
        if FinPeriode.Checked = true then
        // effectif présent en fin de mois
          SQL := SQL + '(PSA_DATESORTIE > "'+ColleZeroDevant(i,2)+'/'+
                       IntToStr(FinMois)+'/'+IntToStr(AnneeDebut)+'" OR '+
                       'PSA_DATESORTIE IS NULL OR PSA_DATESORTIE = "01/01/1900") AND '+
                       'PSA_SUSPENSIONPAIE <> "X"'+WhereEff
        else
        // effectif présent au cours du mois
          SQL := SQL + '(PSA_DATESORTIE IS NULL OR PSA_DATESORTIE = "01/01/1900" OR '+
           'PSA_DATESORTIE >= "'+ColleZeroDevant(i,2)+'/01/'+IntToStr(AnneeDebut)+'") AND '+
           'PSA_SUSPENSIONPAIE <> "X"'+WhereEff;

    QEff := OpenSQL(SQL,True);
    EffPresent [i] :=  QEff.Fields[0].AsInteger;
    Ferme(QEff);
    if (i = MoisDebut) and (VH_Paie.PGDecalage = TRUE) then
    begin
      i := 0;
      AnneeDebut := AnneeDebut+1;
      MoisFin := 11;
    end;
    i := i+1;
  end;


{  if V_PGI.Debug=True then
   PGVisuUnObjet( TobMassCat,'','') ;}

 if (TOBEtat = NIL) then
  begin
    TOBEtat := TOB.Create('Masses salariales à éditer', NIL, -1);
  end;

  Categ := '';
  TobMass := TobMassCat.FindFirst([''],[''],TRUE);

  for i := 0 to TobMassCat.Detail.Count -1 do
  begin
    ConstruireTob;
    TobMass := TobMassCat.FindNext([''],[''],TRUE);
  end;
  TobMassCat.free;
 //debut PT3
{if V_PGI.Debug=True then
   PGVisuUnObjet( TobEtat,'','') ;}
 { StPages := '';
  if GetCheckBoxState('FLISTE') = CbChecked then
    LanceEtatTOB('E','PEV','PMM',TobEtat,True,True,False,Pages,' ','',False,0,StPages)
  else
    LanceEtatTOB('E','PEV','PMM',TobEtat,True,False,False,Pages,' ','',False);
  TobEtat.Free;} //mise en commentaire
  TFQRS1(Ecran).LaTob:= TobEtat;
 //fin PT3
end ;

procedure TOF_PGEVOLUTIONMASSE.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;
procedure TOF_PGEVOLUTIONMASSE.ControlPeriode(Sender: TObject);
begin
PgIFValidPeriode;
end;
procedure TOF_PGEVOLUTIONMASSE.ChangeExercice(Sender: TObject);
var
QPeriode : TQuery;
begin
QPeriode:=OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN),'+
                  'MAX(PEX_ANNEEREFER) FROM PAIEENCOURS '+
  'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN '+
  'WHERE PEX_EXERCICE="'+GetControlText('EDEXERSOC')+'"',TRUE);
if not QPeriode.eof then
  Begin
//  SetControlText('XX_VARIABLEDEB',DateToStr(QPeriode.Fields[0].AsDateTime));
//  SetControlText('XX_VARIABLEFIN',DateToStr(QPeriode.Fields[1].AsDateTime));
  SetControlText('PHC_DATEDEBUT',DateToStr(QPeriode.Fields[0].AsDateTime));
  SetControlText('PHC_DATEFIN',DateToStr(QPeriode.Fields[1].AsDateTime));
//PT1  Annee :=  QPeriode.Fields[2].asstring;
  ExDateDeb := DateToStr(QPeriode.Fields[0].AsDateTime);
  ExDateFin := DateToStr(QPeriode.Fields[1].AsDateTime);

  End
else
  Begin
// d PT1-1
//  SetControlText('XX_VARIABLEDEB',DateToStr(idate1900));
//  SetControlText('XX_VARIABLEFIN',DateToStr(idate1900));
  SetControlText('PHC_DATEDEBUT',DateToStr(idate1900));
  SetControlText('PHC_DATEFIN',DateToStr(idate1900));
// f PT1-1

  ExDateDeb := DateToStr(idate1900);
  ExDateFin := DateToStr(idate1900);
//PT1  Annee := '';
  End;
  Ferme(QPeriode);
end;
procedure TOF_PGEVOLUTIONMASSE.PgIFValidPeriode;
var
  YYD,MMD,JJ,YYF,MMF : WORD;
begin
// d PT1-1
//  if IsValidDate(GetControlText('XX_VARIABLEDEB')) AND
//   IsValidDate(GetControlText('XX_VARIABLEFIN')) Then
  if IsValidDate(GetControlText('PHC_DATEDEBUT')) AND
   IsValidDate(GetControlText('PHC_DATEFIN')) Then
  begin
//  DecodeDate(StrToDate(GetControlText('XX_VARIABLEDEB')),YYD,MMD,JJ);
//  DecodeDate(StrToDate(GetControlText('XX_VARIABLEFIN')),YYF,MMF,JJ);
    DecodeDate(StrToDate(GetControlText('PHC_DATEDEBUT')),YYD,MMD,JJ);
    DecodeDate(StrToDate(GetControlText('PHC_DATEFIN')),YYF,MMF,JJ);

    if (YYF>YYD) and (MMF>=MMD) then
    begin
      PgiBox('La période d''édition ne peut excéder douze mois civiles.','Date Erronée!');
//    SetControlText('XX_VARIABLEFIN',DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('XX_VARIABLEDEB')),11,'M'))));
      SetControlText('PHC_DATEFIN',DateToStr(FinDeMois(PlusDate(StrToDate(GetControlText('PHC_DATEDEBUT')),11,'M'))));
// f PT1-1
   end;
  end;
end;

Procedure TOF_PGEVOLUTIONMASSE.ConstruireTob ();
var
  Ruptlu                : array[1..9] of string;
  Categlu               : string;
  num                   : integer;
  BoolRupt              : array[1..9] of boolean;
begin
  for num := 1 to 9 do
  begin
    BoolRupt[num] := FALSE;
  end;

  if (NbRupt = 0) then
  begin
    Ruptlu[1] := TobMass.getValue(ChampRuptOrig[1]);
    if Ruptlu[1] <> ValeurRupt[1] then BoolRupt[1] := TRUE;
  end
  else
  begin
    for num := 1 to NbRupt do
    begin
      Ruptlu[num] := TobMass.getValue(ChampRuptOrig[num]);
      if Ruptlu[num] <> ValeurRupt[num] then BoolRupt[num] := TRUE;
    end;
  end;
  Categlu := TobMass.getValue('CATEGORIE');

  if ((BoolRupt[1] or
       BoolRupt[2] or
       BoolRupt[3] or
       BoolRupt[4] or
       BoolRupt[5] or
       BoolRupt[6] or
       BoolRupt[7] or
       BoolRupt[8] or
       BoolRupt[9]) or
      (Categ = '')) then
  begin
    T := Tob.Create('Les masses salariales par catégorie',TobEtat,-1);
    if (NbRupt = 0) then  ValeurRupt[1] := TobMass.getValue(ChampRuptOrig[1]);
    for num := 1 to NbRupt do
    begin
      ValeurRupt[num] := TobMass.getValue(ChampRuptOrig[num]);
    end;
    Categ := 'xx';

    for num :=1 to 9 do
    begin
      T.AddChampSupValeur(ChampRupt[num],ValeurRupt[num], False);
      if (NbRupt = 0) then
      begin
        if (num = 1) then
          T.AddChampSupValeur('LIBRUPTURE'+IntToStr(num),
                              TobMass.getValue('ET_LIBELLE'),False);
      end
      else
       if (ChampRuptOrig[num] = 'PHC_ETABLISSEMENT') then
         T.AddChampSupValeur('LIBRUPTURE'+IntToStr(num),
                             TobMass.getValue('ET_LIBELLE'),False)
       else
         T.AddChampSupValeur('LIBRUPTURE'+IntToStr(num),
                             RechDom(Tablette[num],
                             ValeurRupt[num],False));
    end;
// d PT1-1
    T.AddChampSupValeur('PHC_DATEDEBUT',GetControlText('PHC_DATEDEBUT'), False);
    T.AddChampSupValeur('PHC_DATEFIN',GetControlText('PHC_DATEFIN'), False);
// f PT1-1
    T.AddChampSupValeur('EFF01',0,False);  //@@effectif
    T.AddChampSupValeur('EFF02',0,False);  //@@effectif
    T.AddChampSupValeur('EFF03',0,False);  //@@effectif
    T.AddChampSupValeur('EFF04',0,False);  //@@effectif
    T.AddChampSupValeur('EFF05',0,False);  //@@effectif
    T.AddChampSupValeur('EFF06',0,False);  //@@effectif
    T.AddChampSupValeur('EFF07',0,False);  //@@effectif
    T.AddChampSupValeur('EFF08',0,False);  //@@effectif
    T.AddChampSupValeur('EFF09',0,False);  //@@effectif
    T.AddChampSupValeur('EFF10',0,False);  //@@effectif
    T.AddChampSupValeur('EFF11',0,False);  //@@effectif
    T.AddChampSupValeur('EFF12',0,False);  //@@effectif

    T.AddChampSupValeur('EFFP01',EffPresent[1],False);  //@@effectif
    T.AddChampSupValeur('EFFP02',EffPresent[2],False);  //@@effectif
    T.AddChampSupValeur('EFFP03',EffPresent[3],False);  //@@effectif
    T.AddChampSupValeur('EFFP04',EffPresent[4],False);  //@@effectif
    T.AddChampSupValeur('EFFP05',EffPresent[5],False);  //@@effectif
    T.AddChampSupValeur('EFFP06',EffPresent[6],False);  //@@effectif
    T.AddChampSupValeur('EFFP07',EffPresent[7],False);  //@@effectif
    T.AddChampSupValeur('EFFP08',EffPresent[8],False);  //@@effectif
    T.AddChampSupValeur('EFFP09',EffPresent[9],False);  //@@effectif
    T.AddChampSupValeur('EFFP10',EffPresent[10],False);  //@@effectif
    T.AddChampSupValeur('EFFP11',EffPresent[11],False);  //@@effectif
    T.AddChampSupValeur('EFFP12',EffPresent[12],False);  //@@effectif

    T.AddChampSupValeur('MAXMT01',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT01',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT02',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT02',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT03',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT03',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT04',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT04',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT05',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT05',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT06',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT06',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT07',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT07',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT08',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT08',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT09',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT09',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT10',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT10',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT11',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT11',0,False);  //@@effectif
    T.AddChampSupValeur('MAXMT12',0,False);  //@@effectif
    T.AddChampSupValeur('MINMT12',0,False);  //@@effectif

    T.AddChampSupValeur('Ouv01',0.,FALSE);
    T.AddChampSupValeur('Ouv02',0,FALSE);
    T.AddChampSupValeur('Ouv03',0,FALSE);
    T.AddChampSupValeur('Ouv04',0,FALSE);
    T.AddChampSupValeur('Ouv05',0,FALSE);
    T.AddChampSupValeur('Ouv06',0,FALSE);
    T.AddChampSupValeur('Ouv07',0,FALSE);
    T.AddChampSupValeur('Ouv08',0,FALSE);
    T.AddChampSupValeur('Ouv09',0,FALSE);
    T.AddChampSupValeur('Ouv10',0,FALSE);
    T.AddChampSupValeur('Ouv11',0,FALSE);
    T.AddChampSupValeur('Ouv12',0,FALSE);
    T.AddChampSupValeur('Agm01',0,FALSE);
    T.AddChampSupValeur('Agm02',0,FALSE);
    T.AddChampSupValeur('Agm03',0,FALSE);
    T.AddChampSupValeur('Agm04',0,FALSE);
    T.AddChampSupValeur('Agm05',0,FALSE);
    T.AddChampSupValeur('Agm06',0,FALSE);
    T.AddChampSupValeur('Agm07',0,FALSE);
    T.AddChampSupValeur('Agm08',0,FALSE);
    T.AddChampSupValeur('Agm09',0,FALSE);
    T.AddChampSupValeur('Agm10',0,FALSE);
    T.AddChampSupValeur('Agm11',0,FALSE);
    T.AddChampSupValeur('Agm12',0,FALSE);
    T.AddChampSupValeur('Tec01',0,FALSE);
    T.AddChampSupValeur('Tec02',0,FALSE);
    T.AddChampSupValeur('Tec03',0,FALSE);
    T.AddChampSupValeur('Tec04',0,FALSE);
    T.AddChampSupValeur('Tec05',0,FALSE);
    T.AddChampSupValeur('Tec06',0,FALSE);
    T.AddChampSupValeur('Tec07',0,FALSE);
    T.AddChampSupValeur('Tec08',0,FALSE);
    T.AddChampSupValeur('Tec09',0,FALSE);
    T.AddChampSupValeur('Tec10',0,FALSE);
    T.AddChampSupValeur('Tec11',0,FALSE);
    T.AddChampSupValeur('Tec12',0,FALSE);
    T.AddChampSupValeur('Emp01',0,FALSE);
    T.AddChampSupValeur('Emp02',0,FALSE);
    T.AddChampSupValeur('Emp03',0,FALSE);
    T.AddChampSupValeur('Emp04',0,FALSE);
    T.AddChampSupValeur('Emp05',0,FALSE);
    T.AddChampSupValeur('Emp06',0,FALSE);
    T.AddChampSupValeur('Emp07',0,FALSE);
    T.AddChampSupValeur('Emp08',0,FALSE);
    T.AddChampSupValeur('Emp09',0,FALSE);
    T.AddChampSupValeur('Emp10',0,FALSE);
    T.AddChampSupValeur('Emp11',0,FALSE);
    T.AddChampSupValeur('Emp12',0,FALSE);
    T.AddChampSupValeur('Cad01',0,FALSE);
    T.AddChampSupValeur('Cad02',0,FALSE);
    T.AddChampSupValeur('Cad03',0,FALSE);
    T.AddChampSupValeur('Cad04',0,FALSE);
    T.AddChampSupValeur('Cad05',0,FALSE);
    T.AddChampSupValeur('Cad06',0,FALSE);
    T.AddChampSupValeur('Cad07',0,FALSE);
    T.AddChampSupValeur('Cad08',0,FALSE);
    T.AddChampSupValeur('Cad09',0,FALSE);
    T.AddChampSupValeur('Cad10',0,FALSE);
    T.AddChampSupValeur('Cad11',0,FALSE);
    T.AddChampSupValeur('Cad12',0,FALSE);
    T.AddChampSupValeur('Aut01',0,FALSE);
    T.AddChampSupValeur('Aut02',0,FALSE);
    T.AddChampSupValeur('Aut03',0,FALSE);
    T.AddChampSupValeur('Aut04',0,FALSE);
    T.AddChampSupValeur('Aut05',0,FALSE);
    T.AddChampSupValeur('Aut06',0,FALSE);
    T.AddChampSupValeur('Aut07',0,FALSE);
    T.AddChampSupValeur('Aut08',0,FALSE);
    T.AddChampSupValeur('Aut09',0,FALSE);
    T.AddChampSupValeur('Aut10',0,FALSE);
    T.AddChampSupValeur('Aut11',0,FALSE);
    T.AddChampSupValeur('Aut12',0,FALSE);
  end;
  case TobMass.GetValue('MOIS') of
  1 : begin
        T.PutValue ('EFF01',T.GetValue('EFF01')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT01') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT01',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT01') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT01') = 0)) then
          T.PutValue ('MINMT01',TobMass.getValue('MINMT'));
      end;
  2 : begin
        T.PutValue ('EFF02',T.GetValue('EFF02')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT02') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT02',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT02') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT02') = 0)) then
          T.PutValue ('MINMT02',TobMass.getValue('MINMT'));
      end;
  3 : begin
        T.PutValue ('EFF03',T.GetValue('EFF03')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT03') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT03',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT03') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT03') = 0)) then
          T.PutValue ('MINMT03',TobMass.getValue('MINMT'));
      end;
  4 : begin
        T.PutValue ('EFF04',T.GetValue('EFF04')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT04') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT04',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT04') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT04') = 0)) then
          T.PutValue ('MINMT04',TobMass.getValue('MINMT'));
      end;

  5 : begin
        T.PutValue ('EFF05',T.GetValue('EFF05')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT05') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT05',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT05') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT05') = 0)) then
          T.PutValue ('MINMT05',TobMass.getValue('MINMT'));
      end;
  6 : begin
        T.PutValue ('EFF06',T.GetValue('EFF06')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT06') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT06',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT06') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT06') = 0)) then
          T.PutValue ('MINMT06',TobMass.getValue('MINMT'));
      end;
  7 : begin
        T.PutValue ('EFF07',T.GetValue('EFF07')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT07') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT07',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT07') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT07') = 0)) then
          T.PutValue ('MINMT07',TobMass.getValue('MINMT'));
      end;
  8 : begin
        T.PutValue ('EFF08',T.GetValue('EFF08')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT08') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT08',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT08') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT08') = 0)) then
          T.PutValue ('MINMT08',TobMass.getValue('MINMT'));
      end;
  9 : begin
        T.PutValue ('EFF09',T.GetValue('EFF09')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT09') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT09',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT09') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT09') = 0)) then
          T.PutValue ('MINMT09',TobMass.getValue('MINMT'));
      end;
 10 : begin
        T.PutValue ('EFF10',T.GetValue('EFF10')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT10') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT10',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT10') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT10') = 0)) then
          T.PutValue ('MINMT10',TobMass.getValue('MINMT'));
      end;
 11 : begin
        T.PutValue ('EFF11',T.GetValue('EFF11')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT11') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT11',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT11') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT11') = 0)) then
          T.PutValue ('MINMT11',TobMass.getValue('MINMT'));
      end;
 12 : begin
        T.PutValue ('EFF12',T.GetValue('EFF12')+TobMass.getValue('EFF'));
        if (T.GetValue('MAXMT12') <  TobMass.getValue('MAXMT')) then
          T.PutValue ('MAXMT12',TobMass.getValue('MAXMT'));
        if ((T.GetValue('MINMT12') > TobMass.getValue('MINMT')) or
           (T.GetValue('MINMT12') = 0)) then
          T.PutValue ('MINMT12',TobMass.getValue('MINMT'));
      end;
  end;
  if (Categlu = '01') then
  begin
    case TobMass.GetValue('MOIS') of
    1 : T.PutValue ('Ouv01',TobMass.GetValue('MT'));
    2 : T.PutValue ('Ouv02',TobMass.GetValue('MT'));
    3 : T.PutValue ('Ouv03',TobMass.GetValue('MT'));
    4 : T.PutValue ('Ouv04',TobMass.GetValue('MT'));
    5 : T.PutValue ('Ouv05',TobMass.GetValue('MT'));
    6 : T.PutValue ('Ouv06',TobMass.GetValue('MT'));
    7 : T.PutValue ('Ouv07',TobMass.GetValue('MT'));
    8 : T.PutValue ('Ouv08',TobMass.GetValue('MT'));
    9 : T.PutValue ('Ouv09',TobMass.GetValue('MT'));
   10 : T.PutValue ('Ouv10',TobMass.GetValue('MT'));
   11 : T.PutValue ('Ouv11',TobMass.GetValue('MT'));
   12 : T.PutValue ('Ouv12',TobMass.GetValue('MT'));
   end;
  end
  else
    if (Categlu = '04') then
    begin
      case TobMass.GetValue('MOIS') of
      1 : T.PutValue ('Agm01',TobMass.GetValue('MT'));
      2 : T.PutValue ('Agm02',TobMass.GetValue('MT'));
      3 : T.PutValue ('Agm03',TobMass.GetValue('MT'));
      4 : T.PutValue ('Agm04',TobMass.GetValue('MT'));
      5 : T.PutValue ('Agm05',TobMass.GetValue('MT'));
      6 : T.PutValue ('Agm06',TobMass.GetValue('MT'));
      7 : T.PutValue ('Agm07',TobMass.GetValue('MT'));
      8 : T.PutValue ('Agm08',TobMass.GetValue('MT'));
      9 : T.PutValue ('Agm09',TobMass.GetValue('MT'));
     10 : T.PutValue ('Agm10',TobMass.GetValue('MT'));
     11 : T.PutValue ('Agm11',TobMass.GetValue('MT'));
     12 : T.PutValue ('Agm12',TobMass.GetValue('MT'));
      end;
    end
    else
      if (Categlu = '03') then
      begin
        case TobMass.GetValue('MOIS') of
        1 : T.PutValue ('Tec01',TobMass.GetValue('MT'));
        2 : T.PutValue ('Tec02',TobMass.GetValue('MT'));
        3 : T.PutValue ('Tec03',TobMass.GetValue('MT'));
        4 : T.PutValue ('Tec04',TobMass.GetValue('MT'));
        5 : T.PutValue ('Tec05',TobMass.GetValue('MT'));
        6 : T.PutValue ('Tec06',TobMass.GetValue('MT'));
        7 : T.PutValue ('Tec07',TobMass.GetValue('MT'));
        8 : T.PutValue ('Tec08',TobMass.GetValue('MT'));
        9 : T.PutValue ('Tec09',TobMass.GetValue('MT'));
       10 : T.PutValue ('Tec10',TobMass.GetValue('MT'));
       11 : T.PutValue ('Tec11',TobMass.GetValue('MT'));
       12 : T.PutValue ('Tec12',TobMass.GetValue('MT'));
        end;
      end
      else
        if (Categlu = '02') then
        begin
          case TobMass.GetValue('MOIS') of
          1 : T.PutValue ('Emp01',TobMass.GetValue('MT'));
          2 : T.PutValue ('Emp02',TobMass.GetValue('MT'));
          3 : T.PutValue ('Emp03',TobMass.GetValue('MT'));
          4 : T.PutValue ('Emp04',TobMass.GetValue('MT'));
          5 : T.PutValue ('Emp05',TobMass.GetValue('MT'));
          6 : T.PutValue ('Emp06',TobMass.GetValue('MT'));
          7 : T.PutValue ('Emp07',TobMass.GetValue('MT'));
          8 : T.PutValue ('Emp08',TobMass.GetValue('MT'));
          9 : T.PutValue ('Emp09',TobMass.GetValue('MT'));
         10 : T.PutValue ('Emp10',TobMass.GetValue('MT'));
         11 : T.PutValue ('Emp11',TobMass.GetValue('MT'));
         12 : T.PutValue ('Emp12',TobMass.GetValue('MT'));
          end;
        end
        else
          if (Categlu = '12') or (Categlu = '13') then
          begin
            case TobMass.GetValue('MOIS') of
            1 : T.PutValue ('Cad01',TobMass.GetValue('MT'));
            2 : T.PutValue ('Cad02',TobMass.GetValue('MT'));
            3 : T.PutValue ('Cad03',TobMass.GetValue('MT'));
            4 : T.PutValue ('Cad04',TobMass.GetValue('MT'));
            5 : T.PutValue ('Cad05',TobMass.GetValue('MT'));
            6 : T.PutValue ('Cad06',TobMass.GetValue('MT'));
            7 : T.PutValue ('Cad07',TobMass.GetValue('MT'));
            8 : T.PutValue ('Cad08',TobMass.GetValue('MT'));
            9 : T.PutValue ('Cad09',TobMass.GetValue('MT'));
           10 : T.PutValue ('Cad10',TobMass.GetValue('MT'));
           11 : T.PutValue ('Cad11',TobMass.GetValue('MT'));
           12 : T.PutValue ('Cad12',TobMass.GetValue('MT'));
            end;
          end
          else
            if (Categlu > '04') and (Categlu <> '12') and (Categlu <> '13') then
            begin
              case TobMass.GetValue('MOIS') of
              1 : T.PutValue ('Aut01',TobMass.GetValue('MT'));
              2 : T.PutValue ('Aut02',TobMass.GetValue('MT'));
              3 : T.PutValue ('Aut03',TobMass.GetValue('MT'));
              4 : T.PutValue ('Aut04',TobMass.GetValue('MT'));
              5 : T.PutValue ('Aut05',TobMass.GetValue('MT'));
              6 : T.PutValue ('Aut06',TobMass.GetValue('MT'));
              7 : T.PutValue ('Aut07',TobMass.GetValue('MT'));
              8 : T.PutValue ('Aut08',TobMass.GetValue('MT'));
              9 : T.PutValue ('Aut09',TobMass.GetValue('MT'));
             10 : T.PutValue ('Aut10',TobMass.GetValue('MT'));
             11 : T.PutValue ('Aut11',TobMass.GetValue('MT'));
             12 : T.PutValue ('Aut12',TobMass.GetValue('MT'));
            end;
  end;
end;

Initialization
  registerclasses ( [ TOF_PGEVOLUTIONMASSE ] ) ;
end.
