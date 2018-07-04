{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 29/08/2001
Modifié le ... :   /  /
Description .. : Unit de suppression des lignes de moiuvement dans la table
Suite ........ : histosaisrub.
Suite ........ : Cela permet de supprimer les enregistrements provenant
Suite ........ : d'un fichier extérieur générés par exemple par une pointeuse
Suite ........ : ou par un logiciel externe
Mots clefs ... : PAIE
*****************************************************************}
{
PT1  : 08/06/2004 PH V_50 FQ 10359 Rajout critère établissement
PT2  : 06/07/2004 PH V_50 FQ 11333 Suppression des lignes de type MLB associées au type MHE
PT3  : 02/01/2008 FLO V_802 FQ 14755 - Restriction des éléments supprimés à la sélection de l'utilisateur
                                     + affichage d'un message supplémentaire de confirmation
                                     + suppression du code présence si module inaccessible
}
unit UTof_Razenrimport;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,
{$ELSE}

{$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, UTOM, Vierge, P5Util, P5Def, AGLInit, EntPaie,
  PgOutils, Paramdat;

type
  TOF_RAZENRIMPORT = class(TOF)
  public
    procedure OnClose; override;
    procedure OnArgument(Arguments: string); override;
  private
    procedure LanceRaz(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ExecuteSuppression;
  end;

implementation

{ Fonction qui va lancer la suppression des paies sur un exercice
  Avec la possibilité de supprimer aussi les salariés
}

procedure TOF_RAZENRIMPORT.ExecuteSuppression;
var
  DD, FF: TDateTime;
  St, Lib, Valeur : string;   //PT3
  Periodedebut, Periodefin: THEDit;
  Typemvtr, Etab: ThValcombobox;
  i : Integer; //PT3
begin
  Periodedebut := Thedit(getcontrol('PERIODEDEBUT'));
  Periodefin := Thedit(getcontrol('PERIODEFIN'));
  if Periodedebut = nil then exit;
  if PeriodeFin = nil then exit;
  Etab := ThValcombobox(getcontrol('ETAB')); // PT1
  if Etab = nil then exit;
  Typemvtr := ThValcombobox(getcontrol('TYPEMVTR'));
  if Typemvtr = nil then exit;
  if not Isvaliddate(Periodedebut.text) then exit;
  if not Isvaliddate(PeriodeFin.text) then exit;
  DD := strtodate(Periodedebut.text);
  FF := strtodate(Periodefin.text);
  if FF < DD then exit;

  st := 'DELETE FROM HISTOSAISRUB WHERE PSD_DATEDEBUT >= "' +
    usdatetime(DD) + '" AND PSD_DATEFIN <= "' + usdatetime(FF) + '" ';

  //PT3 - Début
  (*  if Typemvtr.value <> '' then
  begin // DEV PT2
    if (Typemvtr.value = 'SRB') or (Typemvtr.value = 'MLB') then
      st := st + ' AND PSD_ORIGINEMVT = "' + Typemvtr.value + '"'
    else // Autre type MHE ou MFP
    begin
      st := st + 'OR (PSD_ORIGINEMVT="MLB" AND EXISTS (SELECT PSD_ORIGINEMVT FROM HISTOSAISRUB HS WHERE HS.PSD_ORIGINEMVT="SRB" AND HS.PSD_DATEDEBUT>="'
        + usdatetime(DD) + '" AND HS.PSD_DATEFIN<="' + usdatetime(FF) + '" AND HS.PSD_SALARIE=PSD_SALARIE))';
    end;
  end // FIN PT2
  else
    st := st + ' AND (PSD_ORIGINEMVT = "MFP" OR PSD_ORIGINEMVT = "MHE" OR ' +
      'PSD_ORIGINEMVT = "MLB" )';*)
  Valeur := Typemvtr.Value;
  If Valeur <> '' Then
  Begin
    St := St + ' AND PSD_ORIGINEMVT="' + Valeur + '"';
    Lib := '#13#10- ' + Typemvtr.Text;
  End
  Else
  Begin
    St := St + ' AND PSD_ORIGINEMVT IN (';
    Lib := '';
    // On se base sur les données de la tablette qui a pu être épurée auparavant
    For i := 0 To Typemvtr.Values.Count-1 Do
    Begin
        If Typemvtr.Values[i] <> '' Then
        Begin
            St := St + '"' + Typemvtr.Values[i] + '",';
            Lib := Lib + '#13#10- ' + Typemvtr.Items[i];
        End;
    End;
    St[Length(St)] := ')';
  End;
  //PT3 - Fin

  // DEB PT1
  if GetControlText('ETAB') <> '' then
    st := st + ' AND PSD_ETABLISSEMENT="' + GetControlText('ETAB') + '"';
  // FIN PT1

  try
    Lib := TraduireMemoire('Attention, vous allez supprimer les mouvements suivants : ')+Lib+'#13#10#13#10'+TraduireMemoire('Etes-vous sûr de vouloir continuer?'); //PT3
    If PGIAsk(Lib, Ecran.Caption) = MrYes Then //PT3
    Begin
        BeginTrans;
        ExecuteSQL(st);
        CommitTrans;
    End;
  except
    Rollback;
    PGIBox(TraduireMemoire('Une erreur est survenue lors de la suppression'), Ecran.Caption);
  end;
end;

procedure TOF_RAZENRIMPORT.LanceRaz(Sender: TObject);
var
  st: string;
begin
  St := TraduireMemoire('Etes-vous sûr de vouloir supprimer les mouvements de la période?');
  if PGIAsk(St, Ecran.Caption) = mrYes then ExecuteSuppression;
end;

procedure TOF_RAZENRIMPORT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_RAZENRIMPORT.OnArgument(Arguments: string);
var
  BtnVal: TToolbarButton97;
  Date: THEdit;
  Typemvtr: ThValcombobox; //PT3
  i : Integer; //PT3
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := LanceRaz;
  if not BlocageMonoPoste(TRUE) then OnClose;

  Date := THEdit(GetControl('PERIODEDEBUT'));
  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;

  Date := THEdit(GetControl('PERIODEFIN'));
  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;


  {VCbxLAnnee := THValComboBox (GetControl ('VCBXANNEE'));
  if VCbxLAnnee <> NIL then VCbxLAnnee.value := '';
  ChbxSal :=  TCheckBox (GetControl ('CHBXDESTSAL'));}

  //PT3 - Début
  If Not VH_Paie.PGModulePresence Then
  Begin
    Typemvtr := THValComboBox(GetControl('TYPEMVTR'));
    if Typemvtr <> Nil Then
    Begin
        For i := 0 To TypeMvtR.Values.Count Do
        Begin
            If TypeMvtR.Values[i] = 'PRE' Then
            Begin
                TypeMvtR.Values.Delete(i);
                TypeMvtR.Items.Delete(i);
                Break;
            End;
        End;
    End;
  End;
  //PT3 - Fin
end;

procedure TOF_RAZENRIMPORT.OnClose;
begin
  DeblocageMonoPoste(TRUE);
end;

initialization
  registerclasses([TOF_RAZENRIMPORT]);
end.

