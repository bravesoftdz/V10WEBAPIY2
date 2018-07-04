unit OuvrFerm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, DBTables, Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, HRichEdt,
  HSysMenu, HDB, HTB97, ColMemo, HPanel, Uiutil, HRichOLE, Hcompte, UtilPgi ;

Procedure OuvreFermeCpte(Lefb : TFichierBase ; Fermer:Boolean ; CliFou : String = '' ) ;

type
  TFOuvrferm = class(TFMul)
    Ouvfer: TCheckBox;
    TG_GENERAL: THLabel;
    NumCpte: THCpteEdit;
    Nature: THValComboBox;
    HNature: THLabel;
    TTc1: THLabel;
    TTd1: THLabel;
    Td1: THCritMaskEdit;
    Td2: THCritMaskEdit;
    TTd2: THLabel;
    Tc1: THCritMaskEdit;
    TTc2: THLabel;
    Tc2: THCritMaskEdit;
    TG_DATECREATION: THLabel;
    TG_DATEDERNMVT: THLabel;
    Mvt1: THCritMaskEdit;
    Dc1: THCritMaskEdit;
    TDc2: THLabel;
    Dc2: THCritMaskEdit;
    TG_DATEDERNMVT2: THLabel;
    Mvt2: THCritMaskEdit;
    TG_ABREGE: THLabel;
    TG_DATEMODIFICATION: THLabel;
    Dm1: THCritMaskEdit;
    TG_DATEMODIFICATION2: THLabel;
    Dm2: THCritMaskEdit;
    LibCpte: TEdit;
    HM: THMsgBox;
    BZgene: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure NatureChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
  private    { Déclarations privées }
    FAvertir : Boolean ;
    Lefb  : TFichierBase ;
    Fermer: Boolean ;
    NbLig : Integer ;
    Pref : String ;
    StCode : String ;
    LaTable : String ;
    LettTable : String ;
    Procedure NommeLesChamps(St : String) ;
    Procedure ModifieComptes ;
    Procedure OuvrirOuFermer ;
    Procedure PositionneHelpContext ;
  public
    CliFou : String ;
  end;

implementation

{$IFNDEF GCGC}
Uses
     CPGeneraux_TOM,
     CPTiers_TOM,
     CPSection_TOM,
     {$IFNDEF CCMP}
     {$IFNDEF CCS3}
     BudSect,Budjal,Budgene,
     {$ENDIF}
     {$ENDIF}
     CPJournal_TOM;
{$ELSE}
Uses Fe_Main ;
{$ENDIF}

{$R *.DFM}

Procedure OuvreFermeCpte(Lefb : TFichierBase ; Fermer:Boolean ; CliFou : String = '' ) ;
Var FOuvrferm : TFOuvrferm ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch'],False,'nrBatch') then Exit ;
FOuvrferm:=TFOuvrferm.Create(Application) ;
FOuvrferm.Q.Manuel:=True ; FOuvrFerm.CliFou:=CliFou ;
Case Lefb of
  fbBudSec1..fbBudSec5: BEGIN
                        FOuvrFerm.Q.Liste:='OUVFERBUDS' ;
                        if Fermer then FOuvrFerm.FNomFiltre:='MULFERBUS'
                                  else FOuvrFerm.FNomFiltre:='MULOUVBUS' ;
                        END ;
  fbBudGen: BEGIN
            FOuvrFerm.Q.Liste:='OUVFERBUDG' ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERBUG'
                      else FOuvrFerm.FNomFiltre:='MULOUVBUG' ;
            END ;
  fbBudJal: BEGIN
            FOuvrFerm.Q.Liste:='OUVFERBUDJ' ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERBUJ'
                      else FOuvrFerm.FNomFiltre:='MULOUVBUJ' ;
            if Not Fermer then FOuvrFerm.HelpContext := 15163000 else FOuvrFerm.HelpContext := 15161000 ;
            END ;
  fbGene  : BEGIN
            FOuvrFerm.Q.Liste:='OUVFERGEN' ;
            if Fermer then FOuvrFerm.HelpContext := 7124000 else FOuvrFerm.HelpContext := 7127000 ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERGEN'
                      else FOuvrFerm.FNomFiltre:='MULOUVGEN' ;
            END ;
  fbJal   : BEGIN
            FOuvrFerm.Q.Liste:='OUVFERJAL' ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERJAL'
                      else FOuvrFerm.FNomFiltre:='MULOUVJAL' ;
            END ;
  fbSect  : BEGIN
            FOuvrFerm.Q.Liste:='OUVFERSEC' ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERSEC'
                      else FOuvrFerm.FNomFiltre:='MULOUVSEC' ;
            END ;
  fbAux   : BEGIN
            FOuvrFerm.Q.Liste:='OUVFERAUX' ;
            if Fermer then FOuvrFerm.FNomFiltre:='MULFERAUX'
                      else FOuvrFerm.FNomFiltre:='MULOUVAUX' ;
            END ;
  else Exit ;
 End ;
FOuvrFerm.Lefb:=Lefb ;
FOuvrFerm.Fermer:=Fermer ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   Try
    FOuvrFerm.ShowModal ;
   Finally
    FOuvrFerm.Free ;
    _Bloqueur('nrBatch',False) ;
   End ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FOuvrFerm,PP) ;
  FOuvrFerm.Show ;
  END ;
END ;


procedure TFOuvrferm.FormShow(Sender: TObject);
Var THNAT : THValComboBox ;
begin
DC1.Text:=StDate1900  ; DC2.Text:=StDate2099 ;
MVT1.Text:=StDate1900 ; MVT2.Text:=StDate2099 ;
DM1.Text:=StDate1900  ; DM2.Text:=StDate2099 ;
PositionneHelpContext ; BZGene.Hint:=HM.Mess[46] ;
if Fermer then OuvFer.State:=cbChecked else OuvFer.State:=cbUnchecked ;
//Bouvrir.Visible:=Fermer ;
Case Lefb of
  fbBudgen:BEGIN
           FListe.MultiFieds := 'BG_BUDGENE;' ;
           NumCpte.ZoomTable:=tzBudgen ; Nature.DataType:='ttSens' ; Nature.ItemIndex:=0 ;
           NumCpte.Name:='BG_BUDGENE' ; Pref:='BG_' ; StCode:='BG_BUDGENE' ; LaTable:='BUDGENE' ;
           Nature.Name:='BG_SENS' ; LettTable:='' ; HNature.Caption:=HM.Mess[26] ;
           if Fermer then Caption:=HM.Mess[4]
                     else BEGIN BOuvrir.Hint:=HM.Mess[41] ; Caption:=HM.Mess[3] ; END ;
           END ;
  fbBudjal:BEGIN
           FListe.MultiFieds := 'BJ_BUDJAL;' ;
           NumCpte.ZoomTable:=tzBudjal ; Nature.Vide:=False ; Nature.DataType:='ttAxe' ;
           NumCpte.Name:='BG_BUDJAL' ; Pref:='BJ_' ; StCode:='BJ_BUDJAL' ; LaTable:='BUDJAL' ;
           If Nature.Values.Count>0 Then Nature.Value:=Nature.Values[0] ; Nature.Name:='BJ_AXE' ; LettTable:='ttBudjal' ;
           TG_GENERAL.Caption:=HM.Mess[43] ; HNature.Caption:=HM.Mess[25] ;
           if Fermer then BEGIN Caption:=HM.Mess[38] ; BOuvrir.Hint:=HM.Mess[44] ; END
                     else BEGIN Caption:=HM.Mess[37] ; BOuvrir.Hint:=HM.Mess[45] ; END ;
           TG_ABREGE.Parent:=PCritere ; TG_ABREGE.Left:=280 ; TG_ABREGE.Top:=17 ;
           LibCpte.Parent:=PCritere ; LibCpte.Left:=336 ; LibCpte.Width:=185 ; LibCpte.Top:=13 ;
           TG_DATECREATION.Parent:=PCritere ; TG_DATECREATION.Left:=280 ; TG_DATECREATION.Top:=53 ;
           Dc1.Parent:=PCritere ; Dc1.Left:=336 ; Dc1.Top:=49 ; Dc1.Width:=68 ;
           Dc2.Parent:=PCritere ; Dc2.Left:=454 ; Dc2.Top:=49 ; Dc2.Width:=68 ;
           TDc2.Parent:=PCritere ; TDc2.Left:=425 ; TDc2.Top:=53 ;
           PComplement.TabVisible:=False ;
           BZGene.Hint:=HM.Mess[47] ;
           END ;
  fbBudSec1..fbBudSec5:BEGIN
                       FListe.MultiFieds := 'BS_AXE;BS_BUDSECT;' ;
                       NumCpte.ZoomTable:=tzBudSec1 ; Nature.Vide:=False ; Nature.DataType:='ttAxe' ;
                       HNature.Caption:=HM.Mess[25] ; NumCpte.Name:='BS_BUDSECT' ; Pref:='BS_' ;
                       StCode:='BS_BUDSECT'; LaTable:='BUDSECT' ;
                       If Nature.Values.Count>0 Then Nature.Value:=Nature.Values[0] ; Nature.Name:='BS_AXE' ;
                       LettTable:='' ;
                       TG_GENERAL.Caption:=HM.Mess[18] ; TG_ABREGE.Caption:=HM.Mess[19] ;
                       if Fermer then BEGIN Caption:=HM.Mess[40] ; BOuvrir.Hint:=HM.Mess[20] ; END
                                 else BEGIN Caption:=HM.Mess[39] ; BOuvrir.Hint:=HM.Mess[21] ; END ;
                       END ;
  fbGene  :BEGIN
           FListe.MultiFieds := 'G_GENERAL;' ;
           NumCpte.ZoomTable:=tzGeneral ; Nature.DataType:='ttNatGene' ;
           NumCpte.Name:='G_GENERAL' ; Pref:='G_' ; StCode:='G_GENERAL'; LaTable:='GENERAUX';
           Nature.Name:='G_NATUREGENE' ; LettTable:='' ;
           if Fermer then Caption:=HM.Mess[6]
                     else BEGIN BOuvrir.Hint:=HM.Mess[41] ; Caption:=HM.Mess[5] ; END ;
           THValComboBox(FindComponent('G_NATUREGENE')).ItemIndex:=0 ;
           END ;
  fbJal   :BEGIN
           FListe.MultiFieds := 'J_JOURNAL;' ;
           NumCpte.ZoomTable:=tzJournal ; Nature.DataType:='ttNatJal' ;
           NumCpte.Name:='J_JOURNAL' ; Pref:='J_' ; StCode:='J_JOURNAL'; LaTable:='JOURNAL' ;
           Nature.Name:='J_NATUREJAL' ; LettTable:='ttJournal' ;
           if Fermer then BEGIN Caption:=HM.Mess[8] ; BOuvrir.Hint:=HM.Mess[22] ; END
                     else BEGIN Caption:=HM.Mess[7] ; BOuvrir.Hint:=HM.Mess[23] ; END ;
           TG_GENERAl.Caption:=HM.Mess[24] ;
           THValComboBox(FindComponent('J_NATUREJAL')).ItemIndex:=0 ;
           BZGene.Hint:=HM.Mess[48] ;
           END ;
  fbSect  :BEGIN
           FListe.MultiFieds := 'S_AXE;S_SECTION;' ;
           NumCpte.ZoomTable:=tzSection ; Nature.Vide:=False ; Nature.DataType:='ttAxe' ;
           HNature.Caption:=HM.Mess[25] ; NumCpte.Name:='S_SECTION' ; Pref:='S_' ;
           StCode:='S_SECTION'; LaTable:='SECTION' ;
           If Nature.Values.Count>0 Then Nature.Value:=Nature.Values[0] ;
           Nature.Name:='S_AXE' ;
           LettTable:='' ;
           if Fermer then BEGIN Caption:=HM.Mess[10] ; BOuvrir.Hint:=HM.Mess[20] ; END
                     else BEGIN Caption:=HM.Mess[9] ; BOuvrir.Hint:=HM.Mess[21] ; END ;
           TG_GENERAL.Caption:=HM.Mess[18] ; TG_ABREGE.Caption:=HM.Mess[19] ;
           BZGene.Hint:=HM.Mess[49] ;
           {$IFDEF CCS3}
           TWinControl(FindComponent('S_AXE')).Visible:=False ; HNature.Visible:=False ;
           {$ENDIF}
           END ;
  fbAux   :BEGIN
           FListe.MultiFieds := 'T_AUXILIAIRE;' ;
           NumCpte.ZoomTable:=tzTiers ; Nature.DataType:='ttNatTiers' ;
           NumCpte.Name:='T_AUXILIAIRE' ; Pref:='T_' ; StCode:='T_AUXILIAIRE'; LaTable:='TIERS';
           Nature.Name:='T_NATUREAUXI' ;  LettTable:='' ;
           if Fermer then BEGIN Caption:=HM.Mess[12] ; BOuvrir.Hint:=HM.Mess[30] ; END
                     else BEGIN Caption:=HM.Mess[11] ; BOuvrir.Hint:=HM.Mess[29] ; END ;
           THNAT:=THValComboBox(FindComponent('T_NATUREAUXI')) ;
           if THNAT<>Nil then
              BEGIN
              THNAT.ItemIndex:=0 ;
              {$IFDEF GCGC}
              if CliFou='CLI' then
                 BEGIN
                 THNAT.Value:='CLI' ; THNAT.Enabled:=False ;
                 if Fermer then Caption:=HM.Mess[53] else Caption:=HM.Mess[52] ;
                 UpdateCaption(Self) ;
                 END else if CliFou='FOU' then
                 BEGIN
                 THNAT.Value:='FOU' ; THNAT.Enabled:=False ;
                 if Fermer then Caption:=HM.Mess[55] else Caption:=HM.Mess[54] ;
                 UpdateCaption(Self) ;
                 END else if CliFou='PRO' then
                 BEGIN
                 THNAT.Value:='PRO' ; THNAT.Enabled:=False ;
                 if Fermer then Caption:=HM.Mess[57] else Caption:=HM.Mess[56] ;
                 UpdateCaption(Self) ;
                 END ;
              {$ENDIF}
              END ;
           END ;
  else Exit ;
 End ;
NommeLesChamps(Pref) ;
  inherited;
UpdateCaption(Self) ;
BChercheClick(Nil) ; FAvertir:=False ;
end;

Procedure TFOuvrferm.NommeLesChamps(St : String) ;
BEGIN
LibCpte.Name:=Pref+'LIBELLE'     ; OuvFer.Name:=Pref+'FERME' ;
Dc1.Name:=Pref+'DATECREATION'    ; Dc2.Name:=Pref+'DATECREATION_' ;
Dm1.Name:=Pref+'DATEMODIF'       ; Dm2.Name:=Pref+'DATEMODIF_' ;
Mvt1.Name:=Pref+'DATEDERNMVT'    ; Mvt2.Name:=Pref+'DATEDERNMVT_' ;
if Not (Lefb in [fbBudgen,fbBudSec1..fbBudSec5,fbBudjal]) then
   BEGIN
   Td1.Name:=Pref+'TOTALDEBIT'  ; Td2.Name:=Pref+'TOTALDEBIT_' ;
   Tc1.Name:=Pref+'TOTALCREDIT' ; Tc2.Name:=Pref+'TOTALCREDIT_' ;
   END else
   BEGIN
   Td1.Visible:=False ; Td2.Visible:=False ; TTd1.Visible:=False ; TTd2.Visible:=False ;
   Tc1.Visible:=False ; Tc2.Visible:=False ; TTc1.Visible:=False ; TTc2.Visible:=False ;
   END ;
END ;

procedure TFOuvrferm.BOuvrirClick(Sender: TObject);
begin
  inherited;
ModifieComptes ;
end;

Procedure TFOuvrferm.OuvrirOuFermer ;
Var StC,StA : String ;
BEGIN
Stc:=Q.FindField(StCode).AsString ;
if Fermer then
   BEGIN
   if lefb=fbSect then
      BEGIN
      StA := Q.FindField('S_AXE').AsString ;
      if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                    ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and S_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
      END else
      if lefb in [fbBudSec1..fbBudSec5] then
         BEGIN
         StA := Q.FindField('BS_AXE').AsString ;
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and BS_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END else
         BEGIN
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="-", '+Pref+'DATEOUVERTURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END ;
   END else
   BEGIN
   if lefb=fbSect then
      BEGIN
      StA := Q.FindField('S_AXE').AsString ;
      if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                    ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and S_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
      END else
      if lefb in [fbBudSec1..fbBudSec5] then
         BEGIN
         StA := Q.FindField('BS_AXE').AsString ;
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'" and BS_AXE="' + StA +'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END else
         BEGIN
         if ExecuteSql('UPDATE '+LaTable+' SET '+Pref+'FERME="X", '+Pref+'DATEFERMETURE="'+UsDateTime(Date)+'", '+
                       ''+Pref+'DATEMODIF="'+UsDateTime(Date)+'" Where '+StCode+'="'+Stc+'"')<>1 then V_PGI.IoError:=oeUnknown ;
         END ;
   END ;
END ;

Procedure TFOuvrferm.ModifieComptes ;
Var Msg,i : Byte ;
BEGIN
NbLig:=Fliste.NbSelected ;
if NbLig<=0 then BEGIN HM.Execute(0,caption,'') ; Exit ; END ;
if Fermer then Msg:=1 else Msg:=2 ;
if (Lefb=fbSect) or (Lefb in [fbBudsec1..fbBudsec5]) then
   BEGIN if Fermer then Msg:=14 else Msg:=15 ; END ;
if (Lefb=fbJal) or (Lefb=fbBudjal) then
   BEGIN if Fermer then Msg:=16 else Msg:=17 ; END ;
if Lefb=fbAux then
   BEGIN if Fermer then Msg:=27 else Msg:=28 ; END ;
if Lefb=fbBudJal then
   BEGIN if Fermer then Msg:=50 else Msg:=51 ; END ;
if HM.Execute(Msg,caption,'')<>mrYes then Exit ;
for i:=0 to NbLig-1 do
  BEGIN
  Fliste.GotoLeBookMark(i) ;
  if Transactions(OuvrirOuFermer,5)<>oeOK then BEGIN MessageAlerte(HM.Mess[13]) ; Break ; END
                                          else FAvertir:=True ;
  END ;

  // MAJ des tablettes - FQ 16036 SBO 21/09/2005
  if Lefb = fbJal then
    begin
    AvertirTable('TTJALSAISIE');
    AvertirTable('TTJALANOUVEAU');
    AvertirTable('TTJALBANQUE');
    AvertirTable('TTJOURNAUX');
    AvertirTable('CPJOURNALIFRS');
    AvertirTable('TTJALSANSECART');
    AvertirTable('TTJALSAISIE');
    AvertirTable('TTJOURNAL');
    end ;

Fliste.ClearSelected ; BChercheClick(Nil) ;
END ;


procedure TFOuvrferm.NatureChange(Sender: TObject);
Var C : THCpteEdit ;
begin
  inherited;
Case Lefb of
     fbSect : BEGIN
              C:=THCpteEdit(FindComponent('S_SECTION')) ;
              Case THValComboBox(Sender).Value[2] of
                   '1' : C.ZoomTable:=tzSection ;
                   '2' : C.ZoomTable:=tzSection2 ;
                   '3' : C.ZoomTable:=tzSection3 ;
                   '4' : C.ZoomTable:=tzSection4 ;
                   '5' : C.ZoomTable:=tzSection5 ;
                 End ;
              END ;
     fbBudSec1..fbBudSec5 :
              BEGIN
              C:=THCpteEdit(FindComponent('BS_BUDSECT')) ;
              Case THValComboBox(Sender).Value[2] of
                   '1' : C.ZoomTable:=tzBudSec1 ;
                   '2' : C.ZoomTable:=tzBudSec2 ;
                   '3' : C.ZoomTable:=tzBudSec3 ;
                   '4' : C.ZoomTable:=tzBudSec4 ;
                   '5' : C.ZoomTable:=tzBudSec5 ;
                 End ;
              END ;
     End ;
end;

procedure TFOuvrferm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
if FAvertir then
   BEGIN
   if Uppercase(LettTable)='TTJOURNAL' then AvertirMultiTable('ttJournal') else
   if Uppercase(LettTable)='TTBUDJAL' then
      BEGIN
      AvertirMultiTable('ttBudjal') ;
      AvertirMultiTable('ttBudjalSais') ;
      END else AvertirTable(LettTable) ;
   END ;
if isInside(Self) then _Bloqueur('nrBatch',False) ;
end;

procedure TFOuvrferm.FListeDblClick(Sender: TObject);
Var St,StA : String ;
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
  Case Lefb of
{$IFNDEF GCGC}
   {$IFNDEF CCMP}
      {$IFNDEF CCS3}
    fbBudgen: BEGIN
              St:=Q.FindField('BG_BUDGENE').AsString ;
              if St<>'' then FicheBudgene(Q,'',St,taConsult,0) ;
              END ;
    fbBudSec1..fbBudSec5: BEGIN
                          St:=Q.FindField('BS_BUDSECT').AsString ;
                          StA:=Q.FindField('BS_AXE').AsString ;
                          if St<>'' then FicheBudSect(Q,StA,St,taConsult,0) ;
                          END ;
    fbBudjal : BEGIN
               St:=Q.FindField('BJ_BUDJAL').AsString ;
               StA:=Q.FindField('BJ_AXE').AsString ;
               if St<>'' then FicheBudjal(Q,StA,St,taConsult,0) ;
               END ;
      {$ENDIF}
    {$ENDIF}
    fbGene  : BEGIN
              St:=Q.FindField('G_GENERAL').AsString ;
              if St<>'' then FicheGene(Q,'',St,taConsult,0) ;
              END ;
    fbJal   : BEGIN
              St:=Q.FindField('J_JOURNAL').AsString ;
              if St<>'' then FicheJournal(Q,'',St,taConsult,0) ;
              END ;
    fbSect  : BEGIN
              St:=Q.FindField('S_SECTION').AsString ;
              StA:=Q.FindField('S_AXE').AsString ;
              if St<>'' then FicheSection(Q,StA,St,taConsult,0) ;
              END ;
{$ENDIF}
    fbAux   : BEGIN
              St:=Q.FindField('T_AUXILIAIRE').AsString ;
              {$IFNDEF GCGC}
              if St<>'' then
                 FicheTiers(Q,'',St,taConsult,1) ;
             {$ELSE}
              if St<>'' then
                 BEGIN
                 if CliFou='CLI' then  V_PGI.DispatchTT (28,taConsult ,St, '','') else V_PGI.DispatchTT (29,taConsult ,St, '','');
                 END ;
             {$ENDIF}
              END ;
    else Exit ;
    End ;
end;

Procedure TFOuvrferm.PositionneHelpContext ;
BEGIN
Case Lefb of
  fbBudgen: HelpContext:=0 ;//if Fermer then HelpContext:= else HelpContext:= ;
  fbBudjal: if Fermer then HelpContext:=15163000 else HelpContext:=15161000 ;
  fbBudSec1..fbBudSec5: HelpContext:=0 ;//if Fermer then HelpContext:= else HelpContext:= ;
  fbGene  : if Fermer then HelpContext:=7176006 else HelpContext:=7176002;
  fbJal   : if Fermer then HelpContext:=7176007 else HelpContext:=7176003;
  fbSect  : if Fermer then HelpContext:=7176008 else HelpContext:=7176004;
{$IFDEF GCGC}
  fbAux   : if Fermer then HelpContext:=110000105 else HelpContext:=110000085 ;
{$ELSE}
  fbAux   : if Fermer then HelpContext:=7176001 else HelpContext:=7176005;
{$ENDIF} End ;
END ;

end.
