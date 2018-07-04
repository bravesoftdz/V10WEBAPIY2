{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 04/04/2003
Modifié le ... :   /  /    
Description .. : Edition bulletin préparatoire
Mots clefs ... : PAIE;BULLETIN
*****************************************************************}
{
PT1 24/09/2003 SB V_42 FQ 10838 Ajout rupture fin de session de paie, clause SQL
PT2 04/05/2004 SB V_50 FQ 11113 Ajout ordre de tri
PT3 24/08/2007 FLO V_8 FQ 14370 Ajout de la date d'entrée du salarié
}
unit UTOFPGEditBullPrepa;

interface

uses Sysutils,Classes,StdCtrls,ComCtrls,
     {$IFDEF EAGLCLIENT}
     eQRS1,
     {$ELSE}
     QRS1,
     {$ENDIF}
     Utof,HCtrls,Hent1,HMsgBox,HQry,Paramsoc;

Type { Record des champs rupture possible }
     ListeRupture = record
       StCoche       : string ;
       StNomChamp    : string ;
       StLibelle     : String ;
     end;
     TOF_PGBULLPREPA = Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; override ;
       private
       StOrderBy : String;
       TriAlpha : Boolean;
       procedure ExitEdit(Sender: TObject);
       procedure AffectPeriodePaie(Sender: TObject);
       procedure AffectDatePaie (DatePaie : TDateTime);
       procedure GenereRupture ( Sender : TObject );
       procedure GenereTriAlpha( Sender : TObject );
     END ;



implementation

uses PgEditOutils2,PgOutils,PGoutils2,EntPaie; 

var PgListeRupture: array[0..9] of ListeRupture =
(
(StCoche : 'CETAB' ; stNomChamp : 'PHB_ETABLISSEMENT'; StLibelle : 'Etablissement' ) ,
(StCoche : 'CN1'   ; stNomChamp : 'PHB_TRAVAILN1'    ; StLibelle : ''              ) ,
(StCoche : 'CN2'   ; stNomChamp : 'PHB_TRAVAILN2'    ; StLibelle : ''              ) ,
(StCoche : 'CN3'   ; stNomChamp : 'PHB_TRAVAILN3'    ; StLibelle : ''              ) ,
(StCoche : 'CN4'   ; stNomChamp : 'PHB_TRAVAILN4'    ; StLibelle : ''              ) ,
(StCoche : 'CN5'   ; stNomChamp : 'PHB_CODESTAT'     ; StLibelle : ''              ) ,
(StCoche : 'CL1'   ; stNomChamp : 'PHB_LIBREPCMB1'   ; StLibelle : ''              ) ,
(StCoche : 'CL2'   ; stNomChamp : 'PHB_LIBREPCMB2'   ; StLibelle : ''              ) ,
(StCoche : 'CL3'   ; stNomChamp : 'PHB_LIBREPCMB3'   ; StLibelle : ''              ) ,
(StCoche : 'CL4'   ; stNomChamp : 'PHB_LIBREPCMB4'   ; StLibelle : ''              )
);


{ TOF_PGBULLPREPA }

procedure TOF_PGBULLPREPA.OnArgument(Arguments: String);
var
  Defaut                                        : ThEdit;
  Combo                                         : THValComboBox;
  Check                                         : TCheckBox;
  Min,Max,ExerPerEncours,DebPer,FinPer          : string;
  i,Num                                         : Integer;
begin
  inherited;
  { Valeur par défaut }
  StOrderBy:='ORDER BY PHB_SALARIE,PHB_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE'; //PT1 { PT2 }
  TriAlpha := False;
  SetControlChecked('CKEURO',VH_Paie.PGTenueEuro);
  SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
  RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
  Defaut:=ThEdit(getcontrol('PHB_SALARIE'));
  If Defaut<>nil then Begin Defaut.text:=Min; Defaut.OnExit:=ExitEdit; End;
  Defaut:=ThEdit(getcontrol('PHB_SALARIE_'));
  If Defaut<>nil then Begin Defaut.text:=Max;  Defaut.OnExit:=ExitEdit; End;
  RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
  Defaut:=ThEdit(getcontrol('PHB_ETABLISSEMENT'));
  If Defaut<>nil then Defaut.text:=Min;
  Defaut:=ThEdit(getcontrol('PHB_ETABLISSEMENT_'));
  If Defaut<>nil then Defaut.text:=Max;
  Combo:=THValComboBox(GetControl('CBMOIS'));
  if combo<>nil then Combo.OnChange:=AffectPeriodePaie;
  Combo:=THValComboBox(GetControl('CBANNEE'));
  if combo<>nil then Combo.OnChange:=AffectPeriodePaie;
  if RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer) then
    AffectDatePaie(StrToDate(DebPer));
  { Rend visible les Org Stat et code Stat en fonction des ParamSoc }
  { Visibilité code organisation }
  SetControlProperty('TBCOMPLEMENT','TabVisible',(VH_Paie.PGNbreStatOrg>0) Or(VH_Paie.PGLibCodeStat<>''));
  SetControlProperty('TBCHAMPLIBRE','Tabvisible',(VH_Paie.PgNbCombo>0));
  For i := 1 to VH_Paie.PGNbreStatOrg do
    begin
    if i = 1 then PgListeRupture[i].StLibelle:=VH_Paie.PGLibelleOrgStat1
    Else if i = 2 then PgListeRupture[i].StLibelle :=VH_Paie.PGLibelleOrgStat2
    Else if i = 3 then PgListeRupture[i].StLibelle :=VH_Paie.PGLibelleOrgStat3
    Else if i = 4 then PgListeRupture[i].StLibelle :=VH_Paie.PGLibelleOrgStat4;
    SetControlText('T'+PgListeRupture[i].StNomChamp,PgListeRupture[i].StLibelle+' de');
    SetControlVisible('T'+PgListeRupture[i].StNomChamp,(PgListeRupture[i].StLibelle<>''));
    SetControlVisible(PgListeRupture[i].StNomChamp,(PgListeRupture[i].StLibelle<>''));
    SetControlVisible('T'+PgListeRupture[i].StNomChamp+'_',(PgListeRupture[i].StLibelle<>''));
    SetControlVisible(PgListeRupture[i].StNomChamp+'_',(PgListeRupture[i].StLibelle<>''));
    SetControlVisible(PgListeRupture[i].StCoche,(PgListeRupture[i].StLibelle<>''));
    end;
  { Visibilité code statistique }
  PgListeRupture[5].StLibelle:=VH_Paie.PGLibCodeStat;
  SetControlText('T'+PgListeRupture[5].StNomChamp,PgListeRupture[5].StLibelle+' de');
  SetControlVisible('T'+PgListeRupture[5].StNomChamp,(PgListeRupture[5].StLibelle<>''));
  SetControlVisible(PgListeRupture[5].StNomChamp,(PgListeRupture[5].StLibelle<>''));
  SetControlVisible('T'+PgListeRupture[5].StNomChamp+'_',(PgListeRupture[5].StLibelle<>''));
  SetControlVisible(PgListeRupture[5].StNomChamp+'_',(PgListeRupture[5].StLibelle<>''));
  SetControlVisible(PgListeRupture[5].StCoche,(PgListeRupture[5].StLibelle<>''));
  { Visibilité champlibre }
  For i := 1 to VH_Paie.PgNbCombo do
    begin
    SetControlProperty('TBCHAMPLIBRE','TabVisible',True);
    num := i + 5;   { SB 18/04/2005 Erreur incrément }
    if i = 1 then PgListeRupture[num].StLibelle:=VH_Paie.PgLibCombo1
    Else if i = 2 then PgListeRupture[num].StLibelle :=VH_Paie.PgLibCombo2
    Else if i = 3 then PgListeRupture[num].StLibelle :=VH_Paie.PgLibCombo3
    Else if i = 4 then PgListeRupture[num].StLibelle :=VH_Paie.PgLibCombo4;
    SetControlText('T'+PgListeRupture[num].StNomChamp,PgListeRupture[num].StLibelle+' de');
    SetControlVisible('T'+PgListeRupture[num].StNomChamp,(PgListeRupture[num].StLibelle<>''));
    SetControlVisible(PgListeRupture[num].StNomChamp,(PgListeRupture[num].StLibelle<>''));
    SetControlVisible('T'+PgListeRupture[num].StNomChamp+'_',(PgListeRupture[num].StLibelle<>''));
    SetControlVisible(PgListeRupture[num].StNomChamp+'_',(PgListeRupture[num].StLibelle<>''));
    SetControlVisible(PgListeRupture[num].StCoche,(PgListeRupture[num].StLibelle<>''));
    end;
  { Evenement ONCLICK }
  Check:=TCheckBox(GetControl('CN1'));
  if Check<>nil then  Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CN2'));
  if Check<>nil then  Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CN3'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CN4'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CN5'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CETAB'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CL1'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CL2'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CL3'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CL4'));
  if Check<>nil then Check.OnClick:=GenereRupture;
  Check:=TCheckBox(GetControl('CALPHA'));
  if Check<>nil then Check.OnClick:=GenereTriAlpha;
end;

procedure TOF_PGBULLPREPA.OnUpdate;
Var
  SQL,Temp,Tempo,Critere : String;
  x                      : integer;
begin
  inherited;
  Temp:=RecupWhereCritere(TPageControl(GetControl('Pages')));
  tempo:=''; critere:='';
  x:=Pos('(',Temp);
  if x>0 then Tempo:=copy(Temp,x,(Length(temp)-5));
  if tempo<>'' then critere:='AND '+Tempo;

  SQL:= 'SELECT PHB_ETABLISSEMENT,PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN,PHB_NATURERUB,'+
        'PHB_RUBRIQUE,PHB_LIBELLE,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM,'+
        'PHB_IMPRIMABLE,PHB_BASEREMIMPRIM,PHB_TAUXREMIMPRIM,PHB_COEFFREMIMPRIM,'+
        'PHB_TRAVAILN1,PHB_TRAVAILN2,PHB_TRAVAILN3,PHB_TRAVAILN4,PHB_ORDREETAT,PHB_SENSBUL,'+
        'PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE,'+   //PT3
        'PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,'+
        'PSA_QUALIFICATION,PSA_COEFFICIENT,PSA_LIBELLEEMPLOI,PSA_PGMODEREGLE,'+
        'PRM_TYPEBASE,PRM_TYPETAUX,PRM_TYPECOEFF,PRM_TYPEMONTANT '+
        'FROM HISTOBULLETIN '+
        'LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE '+
        'LEFT JOIN PAIEENCOURS ON PPU_SALARIE=PHB_SALARIE AND PHB_DATEDEBUT=PPU_DATEDEBUT '+ //PT1
        'AND PHB_DATEFIN=PPU_DATEFIN AND PPU_ETABLISSEMENT=PHB_ETABLISSEMENT '+
        'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE '+
        'WHERE PHB_NATURERUB="AAA" AND PHB_IMPRIMABLE="X" '+
        'AND PHB_RUBRIQUE NOT LIKE "%.%" AND PPU_BULCOMPL="-" '+critere+StOrderBy;           //PT1
  TFQRS1(Ecran).WhereSQL:=SQL;
end;

procedure TOF_PGBULLPREPA.AffectPeriodePaie(Sender: TObject);
var
  Mois,Annee,AnneeOk : string;
  DebPaie,FinPaie    : TDateTime;
begin
  Mois:=GetControlText('CBMOIS');
  Annee:=GetControlText('CBANNEE');
  if (Mois='') or (Annee='') then exit;
  ControlMoisAnneeExer(Mois, RechDom ('PGANNEESOCIALE',Annee,FALSE),AnneeOk );
  if AnneeOk<>'' then
    Begin
    DebPaie:=EncodeDate(StrToInt(AnneeOk),StrToInt(Mois),1);
    FinPaie:=FindeMois(EncodeDate(StrToInt(AnneeOk),StrToInt(Mois),1));
    SetControlText('PHB_DATEDEBUT',DateToStr(DebPaie));
    SetControlText('PHB_DATEFIN',DateToStr(FinPaie));
    End;
end;

procedure TOF_PGBULLPREPA.ExitEdit(Sender: TObject);
var edit : thedit;
begin
  edit:=THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGBULLPREPA.AffectDatePaie(DatePaie: TDateTime);
var YY,MM,JJ : WORD;
    Mois,ExerPerEncours : string;
begin
  if DatePaie>idate1900 then
    Begin
    DecodeDate(DatePaie,YY,MM,JJ);
    Mois:=IntToStr(MM);
    if Length(Mois)=1 then Mois:='0'+Mois;
    SetControlText('CBMOIS',Mois);
    ExerPerEncours:=RendExerciceCorrespondant(DatePaie);
    if (ExerPerEncours='') then
      PgiBox('Aucun exercice ne correspond à cette session de paie','Exercice social erronée');
    SetControlText('CBANNEE',ExerPerEncours);
    AffectPeriodePaie(nil);
    End;
end;

procedure TOF_PGBULLPREPA.GenereRupture(Sender: TObject);
Var
  i       : Integer;
  StAlpha : String;
begin
  if not assigned(Sender) then exit;
  SetControlText('XX_VARIABLE1' ,'');
  SetControlText('XX_RUPTURE1'  ,'');
  if TriAlpha then StAlpha:='PSA_LIBELLE,' else StAlpha:='';
  StOrderBy:='ORDER BY '+StAlpha+'PHB_SALARIE,PHB_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE'; //PT1 { PT2 }
  For i:=0 to 9 Do
    Begin
    if (PgListeRupture[i].StCoche=TCheckBox(Sender).Name) And (TCheckBox(Sender).checked) then
      Begin
      SetControlText('XX_VARIABLE1' ,PGListeRupture[i].StLibelle);
      SetControlText('XX_RUPTURE1'  ,PGListeRupture[i].stNomChamp);
      StOrderBy:='ORDER BY '+PGListeRupture[i].stNomChamp+','+StAlpha+'PHB_SALARIE,PHB_DATEFIN,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE'; //PT1 { PT2 }
      End
    else
      SetControlEnabled(PGListeRupture[i].StCoche,Not(TCheckBox(Sender).checked));
    End;
end;

procedure TOF_PGBULLPREPA.GenereTriAlpha(Sender: TObject);
begin
  TriAlpha:=TCheckBox(Sender).checked;
  if (TriAlpha) and (Pos('PSA_LIBELLE',StOrderBy)<1) then
    StOrderBy:=Copy(StOrderBy,1,Pos('PHB_SALARIE',StOrderBy)-1)+'PSA_LIBELLE,'+Copy(StOrderBy,Pos('PHB_SALARIE',StOrderBy),Length(StOrderBy))
  else
    if (not TriAlpha) and (Pos('PSA_LIBELLE',StOrderBy)>0) then
      SyStem.delete(StOrderBy,Pos('PSA_LIBELLE',StOrderBy),Length('PSA_LIBELLE,'));
  StOrderBy:=Trim(StOrderBy);
end;

Initialization
registerclasses([TOF_PGBULLPREPA]);
end.
