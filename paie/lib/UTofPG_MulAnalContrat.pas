{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 23/08/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du mul des contrats de travail
Mots clefs ... : PAIE;PGCONTRAT
*****************************************************************}
{  PT1  23/08/01 V547 PH non initialisation des dates essai
   PT2  12/08/02 V585 PH Gestion des dates de préavis
   PT3  12/06/03 V_42 JL FQ : 10611 Initialisation date fin contrat a 01/01/1900 pour voir tous les contrats
   PT4  09/02/06 V650 GGS FQ : 12852 Complément à zéros du matricule si < 10
   PT5  28/03/07 V800 GGS FQ 13928 Extraction à date d'arrêté
}
unit UTofPG_MULANALCONTRAT;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, mul, DBGrids, FE_Main,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM,
  AGLInit, ParamDat;

type
  TOF_PGMULANALCONTRAT = class(TOF)
  private
    WW: THEdit;
    FinContratDu, FinContratAu, FinEssaiDu, FinEssaiAu: THEdit; // Bornes de dates pour tests
    //    PT2  12/08/02 V585 PH Gestion des dates de préavis
    FinPreavisDu, FinPreavisAu: THEdit;
    // PT4 complément à zéro des matricules saisis sur - 10 caractères
    Pci_Salarie : THEdit;
    procedure ActiveWhere(Sender: TObject);
    procedure FinContratDuExit(Sender: TObject);
    procedure FinContratAuExit(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure FinEssaiDuExit(Sender: TObject);
    procedure FinEssaiAuExit(Sender: TObject);
    //    PT2  12/08/02 V585 PH Gestion des dates de préavis
    procedure FinPreavisDuExit(Sender: TObject);
    procedure FinPreavisAuExit(Sender: TObject);
    //  PT4 complément à zéro des matricules saisis sur - 10 caractères
    procedure Pci_SalarieExit(Sender: TObject);
//PT5
    procedure CkDatearreteClick(Sender : TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation
uses PgOutils,
     PgOutils2,
     P5Def;

procedure TOF_PGMULANALCONTRAT.ActiveWhere(Sender: TObject);
var
  FinCDu, FinCAu, FinEDu, FinEAu, FinPDu, FinPAu: TDateTime; // Bornes de Dates pour tests
//PT5
  Datearret : TCheckBox;
  arrete :  TDateTime;
  where : String;
  begin
  WW.Text := '';
  FinCDu := StrToDate(FinContratDu.Text);
  FinCAu := StrToDate(FinContratAu.Text);
  if FinEssaiDu <> nil then FinEDu := StrToDate(FinEssaiDu.Text)
  else FinEDu := IDate1900;
  if FinEssaiAu <> nil then FinEAu := StrToDate(FinEssaiAu.Text)
  else FinEAu := IDate1900;
  if (WW <> nil) then
  begin
    WW.Text := ' ( ((PCI_FINCONTRAT >="' + UsDateTime(FinCDu) + '") OR (PCI_FINCONTRAT IS NULL)) AND ((PCI_FINCONTRAT <="' + UsDateTime(FinCAu) + '") OR (PCI_FINCONTRAT IS NULL)))';
    if FinEssaiDu <> nil then
    begin
      if FinEssaiDu.Text <> '' then
      begin
        WW.Text := WW.Text + ' AND ' +
          '( ((PCI_ESSAIFIN >="' + UsDateTime(FinEDu) + '") AND (PCI_RNVESSAIFIN IS NULL)) OR ((PCI_RNVESSAIFIN >="' + UsDateTime(FinEDu) + '") AND (PCI_RNVESSAIFIN IS NOT NULL)) ) AND ' +
          '( ((PCI_ESSAIFIN <="' + UsDateTime(FinEAu) + '") AND (PCI_RNVESSAIFIN IS NULL)) OR ((PCI_RNVESSAIFIN <="' + UsDateTime(FinEAu) + '") AND (PCI_RNVESSAIFIN IS NOT NULL)) )';
      end;
    end;
    //    PT2  12/08/02 V585 PH Gestion des dates de préavis
    FinPDu := 0;
    FinPAu := 0;
    if (FinPreavisDu <> nil) and (Copy(FinPreavisDu.Text, 1, 2) <> '  ') then FinPDu := StrToDate(GetControlText('FINPREAVISDU'));
    if (FinPreavisAu <> nil) and (Copy(FinPreavisAu.Text, 1, 2) <> '  ') then FinPAu := StrToDate(GetControlText('FINPREAVISAU'));
    if FinPDu > 10 then
    begin
      WW.Text := WW.Text + ' AND ' +
        ' ( ((PCI_FINPREAVIS >="' + UsDateTime(FinPDu) + '") OR (PCI_FINPREAVIS IS NULL)) AND ((PCI_FINPREAVIS <="' + UsDateTime(FinPAu) + '") OR (PCI_FINPREAVIS IS NULL)))';
    end;
//PT5
    if TCheckBox(GetControl('DATEARRET')) <> nil then
    begin
      if (GetControlText('Datearret')='X') and (IsValidDate(GetControlText('ARRETE'))) then
      begin
        arrete := StrToDate(GetControlText('arrete'));
        where := 'AND (PCI_DEBUTCONTRAT <= "'+ USDateTime(arrete) +'" AND (PCI_FINCONTRAT >= "'
              + USDateTime(arrete) + '" OR PCI_FINCONTRAT= "' + USDateTime(iDATE1900) + '" OR PCI_FINCONTRAT IS NULL))';
        WW.Text := WW.Text + where;
      end;
    end;
//FIN PT5
  end;
end;

procedure TOF_PGMULANALCONTRAT.OnArgument(Arguments: string);
var
  num: integer;
  datearret : TcheckBox;    //PT5
  arrete : THEdit;          //PT5
begin
  inherited;
  WW := THEdit(GetControl('XX_WHERE'));
  FinContratDu := ThEdit(getcontrol('FCONTRATDU'));
  Pci_Salarie := ThEdit(Getcontrol('PCI_SALARIE'));
    //  PT4 complément à zéro des matricules saisis sur - 10 caractères
  if Pci_Salarie <> nil then Pci_Salarie.OnExit := Pci_SalarieExit;
  if FinContratDu <> nil then FinContratDu.OnExit := FinContratDuExit;
  FinContratAu := ThEdit(getcontrol('FCONTRATAU'));
  if FinContratAu <> nil then FinContratAu.OnExit := FinContratAuExit;
  if (FinContratDu <> nil) and (FinContratAu <> nil) then
  begin

    FinContratDu.Text := DateToStr(IDate1900); //PT3 Idate1900 au lieu de 21916 // '01/01/1960'
    FinContratAu.Text := DateToStr(73045); // '12/31/2099'
    FinContratDu.OnElipsisClick := DateElipsisclick;
    FinContratAu.OnElipsisClick := DateElipsisclick;
  end;
  FinEssaiDu := ThEdit(getcontrol('FESSAIDU'));
  if FinEssaiDu <> nil then FinEssaiDu.OnExit := FinEssaiDuExit;
  FinEssaiAu := ThEdit(getcontrol('FESSAIAU'));
  if FinEssaiAu <> nil then FinEssaiAu.OnExit := FinEssaiAuExit;
  if (FinEssaiDu <> nil) and (FinEssaiAu <> nil) then
  begin
    FinContratDu.OnElipsisClick := DateElipsisclick;
    FinContratAu.OnElipsisClick := DateElipsisclick;
    {  PT1  23/08/01 V547 PH non initialisation des dates essai
       FinEssaiDu.Text := DateToStr (NOW-30);
       FinEssaiAu.Text := DateToStr (NOW+30);

    }
  end;
  //   PT2  12/08/02 V585 PH Gestion des dates de préavis
  FinPreavisDu := ThEdit(getcontrol('FINPREAVISDU'));
  if FinPreavisDu <> nil then FinPreavisDu.OnExit := FinPreavisDuExit;
  FinPreavisAu := ThEdit(getcontrol('FINPREAVISAU'));
  if FinPreavisAu <> nil then FinPreavisAu.OnExit := FinPreavisAuExit;
  if (FinPreavisDu <> nil) and (FinPreavisAu <> nil) then
  begin
    FinPreavisDu.OnElipsisClick := DateElipsisclick;
    FinPreavisAu.OnElipsisClick := DateElipsisclick;
  end;
  // FIN PT2

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
//PT5
  datearret := TCheckBox(GetControl('DATEARRET'));
  If datearret <> Nil Then datearret.OnCLick := ckdatearreteClick;
  arrete := THEdit(GetControl('ARRETE'));
end;

procedure TOF_PGMULANALCONTRAT.FinContratDuExit(Sender: TObject);
begin
  if not IsValidDate(FinContratDu.Text) then
  begin
    PGIBox('La date de début n''est pas valide', 'Gestion des contrats');
    FinContratDu.SetFocus;
  end;
  if StrToDate(FinContratAu.Text) = 0 then FinContratAu.Text := FinContratDu.Text;
end;

procedure TOF_PGMULANALCONTRAT.FinContratAuExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(FinContratAu.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.Caption);
    FinContratAu.SetFocus;
    exit;
  end;
  if StrToDate(FinContratAu.Text) = 0 then FinContratAu.Text := FinContratDu.Text;
  Date2 := StrToDate(FinContratAu.Text);
  Date1 := StrToDate(FinContratDu.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.Caption);
    FinContratDu.SetFocus;
    FinContratAu.Text := '';
    exit;
  end;
end;

procedure TOF_PGMULANALCONTRAT.FinEssaiDuExit(Sender: TObject);
begin
  exit;
  { zone non gérée
  if  NOT IsValidDate(FinEssaiDu.Text) then
   begin
   PGIBox ('La date de début n''est pas valide', 'Gestion des contrats');
   FinEssaiDu.SetFocus ;
   end;
  if StrToDate(FinEssaiAu.Text ) = 0 then FinEssaiAu.Text := FinEssaiDu.Text;
  }
end;

procedure TOF_PGMULANALCONTRAT.FinEssaiAuExit(Sender: TObject);
//  Date1, Date2: TDateTime;
begin
{  exit;
   zone non gérée
  if  NOT IsValidDate(FinEssaiAu.Text) then
   begin
   PGIBox ('La date de fin n''est pas valide', 'Gestion des contrats');
   FinEssaiAu.SetFocus ;
   exit;
   end;
  if StrToDate(FinEssaiAu.Text ) = 0 then FinEssaiAu.Text := FinEssaiDu.Text;
  Date2:=StrToDate(FinEssaiAu.Text);
  Date1:=StrToDate(FinEssaiDu.Text);
  if Date1 > Date2 then
   begin
   PGIBox ('La date de début est supérieure à la date de fin','Gestion des contrats');
   FinEssaiDu.SetFocus;
   FinEssaiAu.Text:='';
   exit;
   end;
  }
end;

procedure TOF_PGMULANALCONTRAT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMULANALCONTRAT.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;
//    PT2  12/08/02 V585 PH Gestion des dates de préavis

procedure TOF_PGMULANALCONTRAT.FinPreavisAuExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(FinPreavisAu.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.Caption);
    FinPreavisAu.SetFocus;
    exit;
  end;
  if StrToDate(FinPreavisAu.Text) = 0 then FinPreavisAu.Text := FinPreavisDu.Text;
  Date2 := StrToDate(FinPreavisAu.Text);
  Date1 := StrToDate(FinPreavisDu.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.Caption);
    FinPreavisDu.SetFocus;
    FinPreavisAu.Text := '';
    exit;
  end;
end;

procedure TOF_PGMULANALCONTRAT.FinPreavisDuExit(Sender: TObject);
begin
  if not IsValidDate(FinPreavisDu.Text) then
  begin
    PGIBox('La date de début n''est pas valide', 'Gestion des contrats');
    FinPreavisDu.SetFocus;
  end;
  if StrToDate(FinPreavisAu.Text) = 0 then FinPreavisAu.Text := FinPreavisDu.Text;
end;
// FIN PT2
//  PT4 complément à zéro des matricules saisis sur - 10 caractères
procedure TOF_PGMULANALCONTRAT.Pci_SalarieExit(Sender: TObject);
var
  Edit : ThEdit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
// FIN PT4
//PT5
procedure TOF_PGMULANALCONTRAT.CkDatearreteClick(Sender : TObject);
var STWhere : String;
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('DATEARRET') = CbChecked then
        begin
                SetControlEnabled('ARRETE',True);
                SetControlEnabled('LABELDATE',True);
                SetControlText('ARRETE',DateToStr(Date));
                THEdit(GetControl('arrete')).OnElipsisClick := DateElipsisclick;
                StWhere := ' PCI_DEBUTCONTRAT <= "'+ USDateTime(date) +'" AND (PCI_FINCONTRAT >= "'
              + USDateTime(date) + '" OR PCI_FINCONTRAT= "' + USDateTime(iDATE1900) + '" OR PCI_FINCONTRAT IS NULL)';
        end
        else
        begin
                SetControlEnabled('ARRETE',False);
                SetControlEnabled('LABELDATE',False);
                StWhere := '';
        end;
        SetControlText('XX_WHERE',StWhere);
end;

initialization
  registerclasses([TOF_PGMULANALCONTRAT]);
end.

