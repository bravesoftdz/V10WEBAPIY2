{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULDADSETAB ()
Mots clefs ... : TOF;PGMULDADSETAB
*****************************************************************}
{
PT1   : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT2   : 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
}
Unit UTOFPG_MULDADSETAB;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
     FE_Main,
{$else}
     eMul,
     uTob,
     MaineAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     PGDADSCommun,
     Hqry,
     PGOutils2,
     AGLInit,
     EntPaie;

Type
  TOF_PGMULDADSETAB = Class (TOF)
    public
    procedure OnUpdate                 ; override;
    procedure OnLoad                   ; override;
    procedure OnArgument (S : String ) ; override;
    procedure OnDisplay                ; override;
    procedure OnClose                  ; override;
    procedure OnCancel                 ; override;

    private
    Annee, Car, Nature : THValComboBox;
    Q_Mul : THQuery; // Query pour changer la liste associee
    WW : THEdit;     // Clause XX_WHERE

{$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
{$ELSE}
    Liste : THGrid;
{$ENDIF}

    procedure ActiveWhere (Sender: TObject);
    procedure Parametrage();
    procedure GrilleDblClick (Sender: TObject);
    procedure CarDateChange (Sender: Tobject);
    procedure NatureChange (Sender: Tobject);
  end;

Implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnUpdate;
begin
Inherited;
Parametrage;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnLoad
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnLoad;
begin
Inherited;
ActiveWhere (Nil);                 //PT1
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnArgument (S : String);
var
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
AnneeE, AnneePrec, ComboExer, MoisE, StPlus : string;
begin
Inherited;
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid (GetControl ('FListe'));
{$ELSE}
Liste:= THGrid (GetControl ('FListe'));
{$ENDIF}
if (Liste <> nil) then
   Liste.OnDblClick:= GrilleDblClick;

Q_Mul:= THQuery (Ecran.FindComponent ('Q'));

Annee:= THValComboBox (GetControl ('ANNEE'));
Nature:= THValComboBox (GetControl ('CNATURE'));
Car:= THValComboBox (GetControl ('CCAR'));
WW:= THEdit (GetControl ('XX_WHERE'));    //PT1

JourJ:= Date;
DecodeDate (JourJ, AnneeA, MoisM, Jour);

if MoisM > 9 then
   AnneePrec:= IntToStr (AnneeA)
else
   AnneePrec:= IntToStr (AnneeA-1);

if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
   begin
   if Annee <> nil then
      begin
      Annee.value:= ComboExer;
      PGExercice := AnneeE;
      end;
   end
else
   PGIBox ('L''exercice '+AnneePrec+' n''existe pas', TFMul (Ecran).Caption);

if Annee <> nil then
   begin
   PGAnnee:= Annee.value;
   Annee.OnChange:= CarDateChange;
   end;

SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));

if Nature <> nil then
   begin
   Nature.OnChange:= NatureChange;
   Nature.Value:= '0151';
   end;

SetControlEnabled ('LNATURE', True);

if Car <> nil then
   begin
   Car.Enabled:= False;
   Car.Visible:= False;
   Car.OnChange:= CarDateChange;
   Car.Value:= 'A00';
   end;

SetControlVisible ('LCAR', False);
//PT1
ActiveWhere(Nil);

StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ET_ETABLISSEMENT', 'Plus', StPlus);
//FIN PT1
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnClose
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnClose;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnDisplay
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnDisplay ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : OnCancel
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.OnCancel ();
begin
Inherited;
end;

//PT1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/11/2007
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.ActiveWhere(Sender: TObject);
begin
if (WW<>nil) then
   WW.Text:= 'ET_FICTIF<>"X"';
end;
//FIN PT1


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.Parametrage();
var
StExer : string;
QExer : TQuery;
begin
PGAnnee:= Annee.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:= OpenSQL (StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice:= QExer.FindField ('PEX_ANNEEREFER').AsString
else
   begin
   PGExercice:= '';
   Ferme (QExer);
   exit;
   end;

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD:= '005';
      if (Car.Value = 'M01') then
         begin
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer:= StrToDate ('01/02/'+PGExercice);
         FinExer:= FinDeMois (StrToDate ('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer:= StrToDate ('01/03/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/04/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer:= StrToDate ('01/05/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/05/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer:= StrToDate ('01/06/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/07/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer:= StrToDate ('01/08/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/08/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer:= StrToDate ('01/09/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/10/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer:= StrToDate ('01/11/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/11/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer:= StrToDate ('01/12/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD:= '004';
      if (Car.Value = 'T01') then
         begin
{PT2
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
}
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
//FIN PT2
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
      PGIBox ('Périodicité interdite', TFMul (Ecran).Caption);
      Car.Value:= 'A00';
      end;
   if Car.Value='A00' then
      begin
      TypeD:= '002';
      DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
      FinExer:= StrToDate ('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD:= '001';
   if NOT QExer.eof then
      begin
      DebExer:= QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
      FinExer:= QExer.FindField ('PEX_DATEFIN').AsDateTime;
      end;
   end;
Ferme (QExer);

SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : Double-click sur la grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.GrilleDblClick(Sender: TObject);
var
Etab, StEtab : String;
reponse : integer;
begin
if ((Q_Mul <> nil) and (Q_Mul.RecordCount = 0)) then
   exit;

{$IFDEF EAGLCLIENT}
TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1) ;
{$ENDIF}

Etab:= Q_Mul.FindField ('ET_ETABLISSEMENT').AsString;

StEtab:= 'SELECT PDS_SALARIE'+
         ' FROM DADSDETAIL WHERE'+
         ' PDS_SALARIE="'+Etab+'" AND'+
         ' PDS_TYPE="'+TypeD+'" AND'+
         ' PDS_ORDRE > 0 AND'+
         ' PDS_EXERCICEDADS = "'+PGExercice+'"';

{$IFNDEF EAGLCLIENT}
TheMulQ:= THQuery (Ecran.FindComponent ('Q'));
{$ELSE}
TheMulQ:= TOB (Ecran.FindComponent ('Q'));
{$ENDIF}
if (PGAnnee<>'') then
   begin
   reponse:= mrYes;
   if (TypeD<>'001') then
      begin
      if (PGExercice<>'2008') then
         reponse:= PGIAsk ('Attention ! Le cahier des charges utilisé#13#10'+
                           'est à la norme DADS-U 2007.#13#10'+
                           'Voulez-vous continuer ?', TFMul (Ecran).Caption);
      end
   else
      begin
      if (PGExercice <> '2007') then
         reponse:= PGIAsk ('Attention ! Le cahier des charges utilisé#13#10'+
                           'est à la norme DADS-U 2007.#13#10'+
                           'Voulez-vous continuer ?', TFMul (Ecran).Caption);
      end;

   if (reponse=mrYes) then
      begin
      if ((ExisteSQL (StEtab)=FALSE) and (PGAnnee<>'')) then
         AGLLanceFiche ('PAY', 'DADS_ETAB', '', Etab+';'+TypeD,
                        'CREATION;'+Etab+';'+TypeD)
      else
         AGLLanceFiche ('PAY', 'DADS_ETAB', '', Etab+';'+TypeD,
                        'MODIFICATION;'+Etab+';'+TypeD);
      end;
   end
else
   PGIBox ('L''année n''est pas valide', TFMul (Ecran).Caption);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : Modification de la date
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.CarDateChange(Sender: TObject);
begin
Parametrage;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSETAB.NatureChange(Sender: TObject);
begin
if (Car <> nil) then
   begin
   if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
      begin
      Car.Enabled := True;
      Car.Visible := True;
      SetControlEnabled('LCAR', True);
      SetControlVisible('LCAR', True);
      end
   else
      begin
      Car.Enabled := False;
      Car.Visible := False;
      Car.Value := 'A00';
      SetControlEnabled('LCAR', False);
      SetControlVisible('LCAR', False);
      end;
   end;

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if (VH_Paie.PGBTP = False) then
      begin
      PGIBox ('Vous ne gérez pas le module BTP', 'DADS-U standard');
      Nature.Value := '0151';
      end;
   end;
Parametrage;
end;



Initialization
registerclasses ([TOF_PGMULDADSETAB]);
end.
