{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/10/2006
Modifié le ... :   /  /    
Description .. : Gestion des pésentations des indicateurs
Mots clefs ... : PAIE;BILAN
*****************************************************************}
{
PT1 13/10/2006 SB V65 FQ 13563 Correction faute orthographe
}
unit UTOFPG_BSINDPRES;

interface

uses
  Classes,Controls,Utof,
  {$IFDEF EAGLCLIENT}
  eMul,
  {$ELSE}
  Mul,
  {$ENDIF}

  Hmsgbox,Htb97,Utob,Hctrls;

type
  TOF_PGMULBSINDPRES = class(TOF)
    procedure OnArgument(Arg: string); override;
    procedure OnLoad;                  override;
  private
     GblPres,GblPred,GblLib : string;
     LectureSeule : Boolean;
     Procedure ClickBOuvrir (Sender : TObject ) ;
     Procedure ClickBSupprimer (Sender : TObject ) ;
  End;

implementation

uses PgOutils,PGOutils2;

{ TOF_PGMULBSINDPRES }

procedure TOF_PGMULBSINDPRES.ClickBOuvrir(Sender: TObject);
Var Tob_mere,Tob_Fille : Tob;
    i,NbCount : integer;
    ToutSelect : Boolean;
begin
NbCount := 0;
If (GetControlText('CBLISTE') = 'DISPO') AND (TFMul(Ecran).FListe.nbSelected = 0) And (TFMul(Ecran).FListe.AllSelected = False) then
  Begin
  PgiBox('Aucune ligne sélectionnée!',Ecran.Caption);  { PT1 }
  Exit;
  End;
ToutSelect := False;
Tob_Mere := Tob.Create('PBSINDDETSEL',nil,-1);

If PgiAsk('Voulez-vous affecter les indicateurs sélectionnés à la présentation?',Ecran.Caption) = MrYes then
  Begin
    if TFMul(Ecran).FListe.AllSelected then
      begin
      ToutSelect := True;
      NbCount :=  TFmul(Ecran).Q.RecordCount;
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Fetchlestous;
      {$ENDIF}
      TFmul(Ecran).Q.First;
      End
    else
      if TFMul(Ecran).FListe.NbSelected > 0 then NbCount := TFMul(Ecran).FListe.NbSelected;
  for i := 0 to NbCount - 1 do
      begin
        {$IFDEF EAGLCLIENT}
        if not ToutSelect then TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
        {$ENDIF}
        if TFMul(Ecran).FListe.NbSelected > 0 then TFMul(Ecran).FListe.GotoLeBookmark(i);
        //Création de la fille à integrer
        Tob_Fille := Tob.Create('PBSINDDETSEL',Tob_Mere,-1);
        Tob_Fille.PutValue('PIL_PREDEFINI', GblPred);
        Tob_Fille.PutValue('PIL_NODOSSIER', PgRendNodossier());
        Tob_Fille.PutValue('PIL_BSPRESENTATION', GblPres);
        Tob_Fille.PutValue('PIL_INDICATEURBS',TFmul(Ecran).Q.FindField('PBI_INDICATEURBS').asstring);
        Tob_Fille.PutValue('PIL_LIBELLE',TFmul(Ecran).Q.FindField('PBI_LIBELLE').asstring);
        if ToutSelect then TFmul(Ecran).Q.Next;
      end;
   If (Assigned(Tob_Mere)) and (Tob_Mere.Detail.count>0) then
        Begin
        Tob_Mere.SetAllModifie(True);
        Tob_Mere.InsertOrUpdateDB(TRUE);
        End;
   TFMul(Ecran).FListe.ClearSelected;
    if ToutSelect then
      begin
      TFMul(Ecran).FListe.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := False;
      end;
   TFMul(Ecran).BCherche.Click;
   End;

end;

procedure TOF_PGMULBSINDPRES.ClickBSupprimer(Sender: TObject);
var i,NbCount : integer;
    St : string;
    ToutSelect : Boolean;
begin
NbCount := 0;
If (GetControlText('CBLISTE') = 'SELECT') AND (TFMul(Ecran).FListe.nbSelected = 0) And (TFMul(Ecran).FListe.AllSelected = False) then
  Begin
  PgiBox('Aucune ligne sélectionnée!',Ecran.Caption); { PT1 }
  Exit;
  End;

If PgiAsk('Voulez-vous retirer les indicateurs sélectionnés de la présentation?',Ecran.Caption) = MrYes then
  Begin
  ToutSelect := False;
  if TFMul(Ecran).FListe.AllSelected then
      begin
      ToutSelect := True;
      NbCount :=  TFmul(Ecran).Q.RecordCount;
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Fetchlestous;
      {$ENDIF}
      TFmul(Ecran).Q.First;
      End
   else
     if TFMul(Ecran).FListe.NbSelected > 0 then NbCount := TFMul(Ecran).FListe.NbSelected;
  for i := 0 to NbCount - 1 do
      begin
        {$IFDEF EAGLCLIENT}
        if not ToutSelect then TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
        {$ENDIF}
        if TFMul(Ecran).FListe.NbSelected > 0 then TFMul(Ecran).FListe.GotoLeBookmark(i);
        //Suppression de l'enregistrement
        st := 'DELETE FROM PBSINDDETSEL '+
        'WHERE PIL_PREDEFINI="'+GblPred+'" '+
        'AND PIL_NODOSSIER="' + PgRendNoDossier() + '" '+
        'AND PIL_BSPRESENTATION="'+GblPres+'" '+
        'AND PIL_INDICATEURBS="'+TFmul(Ecran).Q.FindField('PBI_INDICATEURBS').asstring+'" ';
        ExecuteSQL(st);
        if ToutSelect then TFmul(Ecran).Q.Next;
      end;
   TFMul(Ecran).FListe.ClearSelected;
   if ToutSelect then
    begin
      TFMul(Ecran).FListe.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := False;
    end;
   TFMul(Ecran).BCherche.Click;
  End;



end;

procedure TOF_PGMULBSINDPRES.OnArgument(Arg: string);
Var
Btn : TToolBarButton97;
CEG, STD, DOS : boolean;
begin
  inherited;
GblPres := ReadTokenSt(Arg);
GblPred := ReadTokenSt(Arg);
GblLib  := ReadTokenSt(Arg);
SetControlText('XX_WHERE','PBI_INDICATEURBS IN (SELECT PIL_INDICATEURBS FROM PBSINDDETSEL '+
                          'WHERE ##PIL_PREDEFINI## PIL_BSPRESENTATION="'+GblPres+'")');
SetControlText('CBLISTE','SELECT');

Btn := TToolBarButton97(GetControl('BOuvrir'));
if Assigned(Btn) then Btn.Onclick := ClickBOuvrir;

Btn := TToolBarButton97(GetControl('BSUPPRIMER'));
if Assigned(Btn) then Btn.Onclick := ClickBSupprimer;

 AccesPredefini('TOUS', CEG, STD, DOS);
  if (GblPred = 'CEG') then
  begin
     LectureSeule := (CEG = False);
    //PaieLectureSeule(TFFiche(Ecran)  , (CEG = False));
     SetControlEnabled('BSUPPRIMER', CEG);
     SetControlEnabled('BOUVRIR'   , CEG);
     SetControlEnabled('BTINDPRES' , CEG);
  end
  else
    if (GblPred = 'STD') then
  begin
    LectureSeule := (STD = False);
  //  PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BSUPPRIMER', STD);
    SetControlEnabled('BOUVRIR'   , STD);
    SetControlEnabled('BTINDPRES' , STD);
  end
  else
    if (GblPred = 'DOS') then
  begin
    LectureSeule := (STD = False);
   // PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('BSUPPRIMER', DOS);
    SetControlEnabled('BOUVRIR'   , DOS);
    SetControlEnabled('BTINDPRES' , DOS);
  end;


end;

procedure TOF_PGMULBSINDPRES.OnLoad;
begin
  inherited;
if GetControlText('CBLISTE') = 'SELECT' then
   Begin
   SetControlText('XX_WHERE','PBI_INDICATEURBS IN (SELECT PIL_INDICATEURBS FROM PBSINDDETSEL '+
                             'WHERE ##PIL_PREDEFINI## PIL_BSPRESENTATION="'+GblPres+'")');
   if not LectureSeule then
      Begin
      SetControlEnabled('BSUPPRIMER',True);
      SetControlEnabled('BOUVRIR',False);
      End;
   TFMul(Ecran).Caption := 'Indicateurs sélectionnés pour la présentation : '+GblPres+' '+GblLib;
   End
else
  if GetControlText('CBLISTE') = 'DISPO' then
     Begin
     SetControlText('XX_WHERE','PBI_INDICATEURBS NOT IN (SELECT PIL_INDICATEURBS FROM PBSINDDETSEL '+
                               'WHERE ##PIL_PREDEFINI## PIL_BSPRESENTATION="'+GblPres+'")');
    if not LectureSeule then
      Begin
      SetControlEnabled('BSUPPRIMER',False);
      SetControlEnabled('BOUVRIR',True);
      End;
     TFMul(Ecran).Caption := 'Indicateurs disponibles pour la présentation : '+GblPres+' '+GblLib;
     End;
UpdateCaption(Ecran);

end;



initialization
  registerclasses([TOF_PGMULBSINDPRES]);
end.
