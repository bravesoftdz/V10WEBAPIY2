unit TomDevise;

interface
uses
    Classes, Dialogs, stdctrls, Sysutils, db, dbctrls, dbtables, buttons, Graphics,
    comctrls, extctrls,
    UTOM, UTOB, FichList, HDB, HmsgBox, Hctrls, Ent1, HEnt1, SaisUtil;

const
     msg0  = '0;Devises;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;';
     msg1  = '1;Devises;Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;';
     msg2  = '2;Devises;Vous devez renseigner un code;W;O;O;O;';
     msg3  = '3;Devises;Vous devez renseigner un libellé;W;O;O;O;';
     msg4  = '4;Devises;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;';
     msg5  = 'L''enregistrement est inaccessible';
     msg6  = '6;Devises;Désirez-vous créer les taux de change correspondant à cette devise?;Q;YNC;Y;C;';
     msg7  = '7;Devises;Vous ne pouvez pas supprimer cette devise: des mouvements lui sont associés.;W;O;O;O;';
     msg8  = '8;Devises;Un taux de zéro bloquera la saisie comptable sur cette devise;E;O;O;O;';
     msg9  = '9;Devises;Aucun taux n''a été saisi. Le taux par défaut sera de 1 dans la saisie sur cette devise;E;O;O;O;';
     msg10 = '10;Devises;Vous ne pouvez pas supprimer la devise pivot !;W;O;O;O;';
     msg11 = '11;Devises;Vous ne pouvez pas modifier le nombre de décimal. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg12 = '12;Devises;Vous ne pouvez pas modifier la quotité. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg13 = '13;Devises;Vous ne pouvez pas supprimer cette devise. Il existe des écritures, des pièces ou des tiers avec cette devise !;W;O;O;O;';
     msg14 = '14;Devises;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum !;W;O;O;O;';
     msg15 = '15;Devises;Vous ne pouvez plus modifier la parité. Il existe des écritures avec cette devise !;W;O;O;O;';
     msg16 = '(A saisir sous la forme 1 Euro = xx,xxxxxx';
     msg17 = '17;Devises;Vous ne pouvez pas supprimer la devise pivot ou nationale;W;O;O;O;';

     msc0  = '0;Chancellerie;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;';
     msc1  = '1;Chancellerie;Confirmez-vous la suppression de l''enregistrement?;Q;YNC;N;C;';
     msc2  = '2;Chancellerie;Vous devez renseigner un code.;W;O;O;O;';
     msc3  = '3;Chancellerie;Vous devez renseigner un libellé.;W;O;O;O;';
     msc4  = '4;Chancellerie;La date de valeur que vous avez saisie est déjà renseignée.;W;O;O;O;';
     msc5  = '5;Chancellerie;Vous devez renseigner une date de valeur.;W;O;O;O;';
     msc6  = '6;Chancellerie;Vous devez renseigner un taux de change pour chaque date de valeur.;W;O;O;O;';
     msc7  = 'L''enregistrement est inaccessible';
     msc8  = '8;Chancellerie;L''enregistrement n''a pas pu être sauvegardé. La date est vide ou elle existe déjà.;W;O;O;O;';
     msc9  = '9;Chancellerie;Vous devez renseigner un taux de change supérieur à zéro.;W;O;O;O;';
     msc10 = 'Valable jusqu''au';
     msc11 = 'Valable à partir du';
     msc12 = 'par rapport à la devise pivot';
     msc13 = 'par rapport à l''Euro';
     msc14 = '14;Chancellerie;La date saisie est supérieure à la date d''entrée en vigueur de l''Euro.;W;O;O;O;';
     msc15 = '15;Chancellerie;La date saisie est inférieure à la date d''entrée en vigueur de l''Euro.;W;O;O;O;';
     msc16 = '1 Euro = xx,xxx $';
     msc17 = '1 $ = xx,xxx Euro';


type
    CptEtLibelle = record
       compte  : string;
       libelle : string;
    end;

    TOM_DEVISE = class(TOM)
        procedure OnUpdateRecord ; override ;
        procedure OnAfterUpdateRecord ; override ;
        procedure OnLoadRecord ; override ;
        procedure OnChangeField (F : TField); override ;
        procedure OnNewRecord  ; override ;
        procedure OnDeleteRecord ; override ;
        procedure OnArgument(Arguments : String ); override ;
        procedure MonnaieInClick(Sender: TObject);
        procedure ChancellerieClick(Sender: TObject);
        procedure ChancellerieOutClick(Sender: TObject);
        procedure BValiderClick(Sender: TObject);
        procedure OnClose; override ;
    private
      OldDec,OldQuot : Integer ;
      OldParite      : double ;
      Favertir,AvecMvt, bFromAss : Boolean ;
      MemoDev  : String3 ;

      DEVPIVOT     : TCheckBox;
      D_MONNAIEIN  : TCheckBox;
      D_FONGIBLE   : TCheckBox;

      D_CPTLETTRDEBIT,
      D_CPTLETTRCREDIT,
      D_CPTPROVDEBIT,
      D_CPTPROVCREDIT : THEdit;

      TD_CPTLETTRDEBIT,
      TD_CPTLETTRCREDIT,
      TD_CPTPROVDEBIT,
      TD_CPTPROVCREDIT :  THLabel;

      TTD_CPTLETTRDEBIT,
      TTD_CPTLETTRCREDIT,
      TTD_CPTPROVDEBIT,
      TTD_CPTPROVCREDIT: THLabel;

      D_QUOTITE     : THEdit;
      D_DECIMALE    : THSpinEdit;
      D_PARITEEURO  : THEdit;
      DateDebutEuro :  THLabel;

      TSREGUL,
      TSEURO : TTabSheet;
      BChancel,
      BChancelOut,
      BValider : TSpeedButton;

      function  GetCompteLibelle(cpt: string): CptEtLibelle;
      procedure PosDevPivot;
      function  VerifPariteEuro : boolean;
      Procedure EstMouvementer(St : String);
      Function  MouvementDepuisEuro(St : String) : boolean ;
      procedure D_PARITEEUROChange;
    end;

    TOM_CHANCELL = class(TOM)
        procedure OnArgument(Arguments : String ); override ;
        procedure OnLoadRecord ; override ;
        procedure OnChangeField (F : TField); override ;
        procedure RGSensClick(Sender: TObject);
        procedure BPurgeClick(Sender: TObject);
        procedure BImprimerClick(Sender: TObject);
    private
        Action    : TActionFiche;
        DateD, DateF: string;
        PDevise   : TPanel;
        FListe    : THDBGrid;
        RGSens    : THRadioGroup;
        H_Devise  : THValComboBox;
        SurEuro, MonnaieIn : boolean;
        TValable, TQuotite, TValeurQ : THLabel;

        BImprimer : TSpeedButton;
        BPurge    : TSpeedButton;

        procedure CodeDeviseChange;
        procedure GereSens(ForceCol : boolean );
        procedure PurgeDates(DatesSup : TDateTime );
        function  QuotiteChange : Integer ;
    end;

implementation
uses
    FE_Main, Hcompte, paramsoc, Purge, PrintDBG;


{ TOM_DEVISE }

procedure TOM_DEVISE.OnArgument(Arguments: String);
var s, s1: string;
begin
  inherited;
      Favertir := false;
      bFromAss := false;
      s := uppercase(Arguments);
      while s <> '' do begin
            s1 := ReadTokenSt(s);
            bFromAss := s1 = 'ASSISTANT=TRUE';
            if bFromAss then
               break;
      end;
      DEVPIVOT     := TCheckBox(GetControl('DEVPIVOT'));
      D_MONNAIEIN  := TCheckBox(GetControl('D_MONNAIEIN'));
      if D_MONNAIEIN <> nil then
         D_MONNAIEIN.OnClick := MonnaieInClick;
      D_FONGIBLE  := TCheckBox(GetControl('D_FONGIBLE'));
      D_CPTLETTRDEBIT    := THEdit(GetControl('D_CPTLETTRDEBIT'));
      D_CPTLETTRCREDIT   := THEdit(GetControl('D_CPTLETTRCREDIT'));
      D_CPTPROVDEBIT     := THEdit(GetControl('D_CPTPROVDEBIT'));
      D_CPTPROVCREDIT    := THEdit(GetControl('D_CPTPROVCREDIT'));

      TD_CPTLETTRDEBIT   := THLabel(GetControl('TD_CPTLETTRDEBIT'));
      TD_CPTLETTRCREDIT  := THLabel(GetControl('TD_CPTLETTRCREDIT'));
      TD_CPTPROVDEBIT    := THLabel(GetControl('TD_CPTPROVDEBIT'));
      TD_CPTPROVCREDIT   := THLabel(GetControl('TD_CPTPROVCREDIT'));

      TTD_CPTLETTRDEBIT  := THLabel(GetControl('TTD_CPTLETTRDEBIT'));
      TTD_CPTLETTRCREDIT := THLabel(GetControl('TTD_CPTLETTRCREDIT'));
      TTD_CPTPROVDEBIT   := THLabel(GetControl('TTD_CPTPROVDEBIT'));
      TTD_CPTPROVCREDIT  := THLabel(GetControl('TTD_CPTPROVCREDIT'));

      D_QUOTITE          := THEdit(GetControl('D_QUOTITE'));
      D_DECIMALE         := THSpinEdit(GetControl('D_DECIMALE'));
      D_PARITEEURO       := THEdit(GetControl('D_PARITEEURO'));
      DateDebutEuro      := THLabel(GetControl('DATEDEBUTEURO'));

      TSREGUL   := TTabSheet(GetControl('TSREGUL'));
      TSEURO    := TTabSheet(GetControl('TSEURO'));

      BChancel := TSpeedButton(GetControl('BCHANCEL'));
      if BChancel <> nil then
         BChancel.OnClick := ChancellerieClick;
      BChancelOut := TSpeedButton(GetControl('BCHANCELOUT'));
      if BChancelOut <> nil then
         BChancelOut.OnClick := ChancellerieOutClick;
      BValider := TSpeedButton(GetControl('BVALIDER'));
      if BValider <> nil then
         BValider.OnClick := BValiderClick;
end;

function TOM_DEVISE.GetCompteLibelle(cpt: string): CptEtLibelle;
var
   s, sql: string;
   q: TQuery;
begin
   s := trim(cpt);
   if s = '' then begin
      result.compte := '';
      result.libelle := '';
      exit;
   end;
   sql := 'SELECT G_GENERAL, G_LIBELLE From GENERAUX ' +
          'WHERE G_GENERAL Like "' + s +'%" ' + RecupWhere(tzGNonCollectif);
   q := OpenSQL(sql, true);
   if not q.eof then begin
      result.compte := q.fieldbyname('G_GENERAL').asstring;
      result.libelle := q.fieldbyname('G_LIBELLE').asstring;
   end;
   Ferme(q);
end;


procedure TOM_DEVISE.OnChangeField(F: TField);
var
   cptLibelle: CptEtLibelle;
begin
  if ds.state = dsBrowse then
     exit;
  if F.FieldName = 'D_FONGIBLE' then
     PosDevPivot;
  if F.FieldName = 'D_PARITEEURO' then
     D_PARITEEUROChange;

  BChancel.Enabled    := not (ds.State in [dsEdit,dsInsert]);
  BChancelOut.Enabled := not (ds.State in [dsEdit,dsInsert]);
  if GetField('D_DEVISE') = V_PGI.DevisePivot then
     BChancel.Enabled := False;
  if VH^.TenueEuro and (GetField('D_DEVISE') = V_PGI.DevisePivot) then
     BChancelOut.Enabled := False;

  if (F.FieldName = 'D_CPTLETTRDEBIT') or (F.FieldName = 'D_CPTLETTRCREDIT') or
     (F.FieldName = 'D_CPTPROVDEBIT') or (F.FieldName = 'D_CPTPROVCREDIT') then begin
      cptLibelle := GetCompteLibelle(GetField(F.FieldName));
      if F.Asstring <> cptLibelle.Compte then
            SetField(F.FieldName, cptLibelle.Compte);
      if (F.FieldName = 'D_CPTLETTRDEBIT') then
            TTD_CPTLETTRDEBIT.Caption := cptLibelle.libelle
      else if (F.FieldName = 'D_CPTLETTRCREDIT') then
            TTD_CPTLETTRCREDIT.Caption := cptLibelle.libelle
      else if (F.FieldName = 'D_CPTPROVDEBIT') then
            TTD_CPTPROVDEBIT.Caption := cptLibelle.libelle
      else
            TTD_CPTPROVCREDIT.Caption := cptLibelle.libelle;
  end;

  if (F.FieldName = 'D_DECIMALE') and (ds.State in [dsEdit]) and (OldDec <> GetField('D_DECIMALE')) then begin
     EstMouvementer(GetField('D_DEVISE'));
     if AvecMvt then begin
          HShowMessage(msg11, '', '');
          SetField('D_DECIMALE', OldDec);
          Abort;
          Exit;
     end;
  end;

  if (F.FieldName = 'D_QUOTITE') and (ds.State in [dsEdit]) and (OldQuot <> GetField('D_QUOTITE')) then begin
     EstMouvementer(GetField('D_DEVISE'));
     if AvecMvt then begin
          HShowMessage(msg12, '', '');
          SetField('D_QUOTITE', OldQuot);
          Abort;
          Exit;
     end;
  end;

  if (F.FieldName = 'D_PARITEEURO') and (not VerifPariteEuro) then begin
     HShowMessage(msg14, '', '');
     Abort;
     Exit;
  end;

  if (F.FieldName = 'D_PARITEEURO') and (ds.State in [dsEdit]) and (OldParite <> GetField('D_PARITEEURO')) and (Not V_PGI.SAV) then begin
     if MouvementDepuisEuro(GetField('D_DEVISE')) then begin
          HShowMessage(msg15, '', '');
          SetField('D_PARITEEURO', OldParite);
          Abort;
          Exit;
     end;
  end;

{$IFDEF SPEC302}
     if (DEVPIVOT.Visible) and (DEVPIVOT.Checked) and (V_PGI.DevisePivot='') and bFromAss then
     begin
          ExecuteSQL('UPDATE SOCIETE SET SO_DEVISEPRINC="' + GetField('D_DEVISE') + '"');
          V_PGI.DevisePivot := GetField('D_DEVISE');
     end;
     if GetField('D_DEVISE') = V_PGI.DevisePivot then
     begin
          V_PGI.OkDecV := GetField('D_DECIMALE');
          ExecuteSql('Update SOCIETE SET SO_DECVALEUR='+IntToStr(V_PGI.OkDecV)+', '+
                     'SO_TAUXEURO='+StrfPoint(GetField('D_PARITEEURO'))+' '+
                     'Where SO_SOCIETE="'+V_PGI.CodeSociete+'" AND SO_DEVISEPRINC="'+GetField('D_Devise')+'"');
     end;
{$ELSE}
     if (DEVPIVOT.Visible) and (DEVPIVOT.Checked) and (V_PGI.DevisePivot='') and bFromAss then
     begin
        SetParamSoc('SO_DEVISEPRINC',GetField('D_DEVISE'));
        V_PGI.DevisePivot := GetField('D_DEVISE');
     end;
     if GetField('D_Devise') = V_PGI.DevisePivot then
     begin
          V_PGI.OkDecV := GetField('D_DECIMALE');
          SetParamSoc('SO_DECVALEUR', V_PGI.OkDecV);
          SetParamSoc('SO_TAUXEURO', GetField('D_PARITEEURO'));
     end;
{$ENDIF}
  inherited;
end;

function TOM_DEVISE.VerifPariteEuro : boolean ;
var StParite : String ;
    Parite   : Integer ;
begin
     Result:=True ;
     StParite:=Trim(D_PARITEEURO.Text) ;
     if Pos(V_PGI.SepDecimal,StParite)<>0 then
        Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1) ;
     Parite:=StrToInt(StParite) ;
     if Parite<=0 then begin
        Result:=False;
        Exit;
     end;
     StParite:=IntToStr(Parite) ;
     if Length(StParite)>6 then
        Result:=False ;
end;


procedure TOM_DEVISE.OnDeleteRecord;
begin
  if VH^.TenueEuro and (GetField('D_DEVISE') = V_PGI.DeviseFongible) then  begin
     HShowMessage(msg17, '', '');
     Exit;
  end;
  try
    SourisSablier;
    EstMouvementer(GetField('D_DEVISE'));
    if AvecMvt then begin
       HShowMessage(msg13, '', '');
       Exit;
    end;
    ExecuteSql('Delete From CHANCELL Where H_DEVISE="'+GetField('D_DEVISE')+'"');
    inherited;
    Favertir := True;
  finally
    SourisNormale;
  end;
end;

procedure TOM_DEVISE.OnLoadRecord;
var
   cpt: string;
begin
  inherited;
  OldDec    := GetField('D_DECIMALE');
  OldQuot   := GetField('D_QUOTITE');
  OldParite := GetField('D_PARITEEURO');

  TTD_CPTLETTRDEBIT.Caption := GetCompteLibelle(GetField('D_CPTLETTRDEBIT')).libelle;
  TTD_CPTLETTRCREDIT.Caption := GetCompteLibelle(GetField('D_CPTLETTRCREDIT')).libelle;
  TTD_CPTPROVDEBIT.Caption := GetCompteLibelle(GetField('D_CPTPROVDEBIT')).libelle;
  TTD_CPTPROVCREDIT.Caption := GetCompteLibelle(GetField('D_CPTPROVCREDIT')).libelle;

  if (D_MONNAIEIN.Checked) or (GetField('D_DEVISE') = V_PGI.DevisePivot) then
  begin
      D_CPTLETTRDEBIT.Enabled := false;
      D_CPTLETTRCREDIT.Enabled := false;
      D_CPTPROVDEBIT.Enabled := false;
      D_CPTPROVCREDIT.Enabled := false;
      TD_CPTLETTRDEBIT.Enabled := false;
      TD_CPTLETTRCREDIT.Enabled := false;
      TD_CPTPROVDEBIT.Enabled := false;
      TD_CPTPROVCREDIT.Enabled := false;

      D_CPTLETTRDEBIT.Color := clBtnFace;
      D_CPTLETTRCREDIT.Color := clBtnFace;
      D_CPTPROVDEBIT.Color := clBtnFace;
      D_CPTPROVCREDIT.Color := clBtnFace;
  end else
  begin
      D_CPTLETTRDEBIT.Enabled := true;
      D_CPTLETTRCREDIT.Enabled := true;
      D_CPTPROVDEBIT.Enabled := true;
      D_CPTPROVCREDIT.Enabled := true;
      TD_CPTLETTRDEBIT.Enabled := true;
      TD_CPTLETTRCREDIT.Enabled := true;
      TD_CPTPROVDEBIT.Enabled := true;
      TD_CPTPROVCREDIT.Enabled := true;

      D_CPTLETTRDEBIT.Color := clWindow;
      D_CPTLETTRCREDIT.Color := clWindow;
      D_CPTPROVDEBIT.Color := clWindow;
      D_CPTPROVCREDIT.Color := clWindow;

      BChancel.Enabled := True ;
  end;
  MonnaieInClick(nil);
  if (GetField('D_FONGIBLE') = 'X') and (VH^.TenueEuro) then begin
     D_PARITEEURO.Enabled := False;
     D_FONGIBLE.Enabled := False ;
  end;
  THLabel(GetControl('HL3')).Visible := D_PARITEEURO.Enabled;
  THLabel(GetControl('HL3')).Caption := msg16 + ' ' +
                                        GetField('D_SYMBOLE') + ' ' + GetField('D_DEVISE') + ')' ;
  BChancelOut.Visible := Not (D_MONNAIEIN.Checked);
  BChancel.Enabled    := not (ds.State in [dsEdit,dsInsert]);
  BChancelOut.Enabled := not (ds.State in [dsEdit,dsInsert]);
  TSEURO.TabVisible := not (VH^.TenueEuro and (GetField('D_DEVISE') = V_PGI.DevisePivot));
  PosDevPivot ;
end;

procedure TOM_DEVISE.OnNewRecord;
begin
  inherited;
  SetField('D_DECIMALE', 2) ;
  SetField('D_QUOTITE', 1);
  SetField('D_FONGIBLE', '-');
  SetField('D_MONNAIEIN', '-');
  SetField('D_PARITEEURO', 1);
end;

procedure TOM_DEVISE.OnUpdateRecord;
begin
  inherited;

end;

procedure TOM_DEVISE.MonnaieInClick(Sender: TObject);
var
   b: boolean;
begin
  b := ExisteSql('Select D_DEVISE from DEVISE Where D_FONGIBLE="X" AND D_DEVISE<>"'+ GetField('D_DEVISE') +'"');
  D_FONGIBLE.Enabled := D_MONNAIEIN.Checked and (not b);
  if not (D_MONNAIEIN.Checked) then
  begin
       if D_FONGIBLE.Checked then
          D_FONGIBLE.Checked := False ;
       DateDebutEuro.Caption := '';
  end else
  begin
       if not D_FONGIBLE.Enabled and D_FONGIBLE.Checked then
          D_FONGIBLE.Checked:=False ;
       DateDebutEuro.Caption:=DateToStr(V_PGI.DateDebutEuro);
  end;
  D_PARITEEURO.Enabled := ((D_MONNAIEIN.Checked)) and (V_PGI.DateDebutEuro >= EncodeDate(1999,01,01))
                       and (V_PGI.DateDebutEuro < EncodeDate(1999,12,31));
  if V_PGI.SAV then
     D_PARITEEURO.Enabled := True;

  DateDebutEuro.Enabled := D_MONNAIEIN.Checked;
  THLabel(GetControl('TD_PARITEEURO')).Enabled := D_MONNAIEIN.Checked;
  THLabel(GetControl('HL4')).Enabled := D_MONNAIEIN.Checked;
end;

procedure TOM_DEVISE.PosDevPivot;
begin
     DEVPIVOT.Visible := (D_FONGIBLE.Checked) and (V_PGI.DevisePivot='') and bFromAss;
     if Not DEVPIVOT.Visible then DEVPIVOT.Checked := False ;
end;


procedure TOM_DEVISE.ChancellerieClick(Sender: TObject);
begin
     if FAvertir then AvertirTable('ttDevise') ;
     AGLLanceFiche('CP','CHANCEL',GetField('D_DEVISE'),'','ACTION=MODIFICATION;SUREURO=FALSE');
//     FicheChancel(D_DEVISE.Text,False,0,Mode,False) ;
end;

procedure TOM_DEVISE.ChancellerieOutClick(Sender: TObject);
begin
     if FAvertir then begin
        AvertirTable('ttDevise');
        AvertirTable('ttDeviseEtat');
     end;
     AGLLanceFiche('CP','CHANCEL',GetField('D_DEVISE'),'','ACTION=MODIFICATION;SUREURO=TRUE');
//     FicheChancel(D_DEVISE.Text,False,0,Mode,True) ;
end;

procedure TOM_DEVISE.BValiderClick(Sender: TObject);
var
   QLoc : TQuery;
begin
   if ds.State in [dsInsert] then begin
      if (GetField('D_DEVISE') <> V_PGI.DevisePivot) and (not D_FONGIBLE.Checked) then begin
         AvertirTable('ttDevise') ;
         if not D_MONNAIEIN.Checked then begin
//            FicheChancel(D_DEVISE.Text,False,0,taCreat,True) ;
            AGLLanceFiche('CP','CHANCEL',GetField('D_DEVISE'),'','ACTION=CREATION;SUREURO=TRUE');
            QLoc:=OpenSql('Select H_TAUXREEL From CHANCELL Where H_DEVISE="'+GetField('D_DEVISE')+'"',True);
            if QLoc.Eof then
               HShowMessage(msg9, '', '')
            else if(QLoc.Fields[0].AsFloat=0) then
               HShowMessage(msg8, '', '');
            Ferme(QLoc) ;
         end;
      end;
      MemoDev := GetField('D_DEVISE');
   end;
end;

procedure TOM_DEVISE.EstMouvementer(St: String);
begin
     AvecMvt:=ExisteSql('Select D_DEVISE from DEVISE Where D_DEVISE="'+St+'" '+
                        'AND (Exists(Select E_DEVISE from ECRITURE Where E_DEVISE="'+St+'"))') ;

end;

function TOM_DEVISE.MouvementDepuisEuro(St: String): boolean;
Var QQ : TQuery ;
begin
     Result:=False ;
     if Not EstMonnaieIn(St) then Exit ;
     QQ:=OpenSQL('Select E_DEVISE from ECRITURE Where E_DATECOMPTABLE>="'+UsDateTime(V_PGI.DateDebutEuro)+'" AND E_EXERCICE="'+QuelExoDT(V_PGI.DateDebutEuro)+'" AND E_DEVISE="'+St+'"',True) ;
     if Not QQ.EOF then Result:=True ;
     Ferme(QQ) ;
end;

procedure TOM_DEVISE.OnAfterUpdateRecord;
begin
  inherited;
  Favertir := True;
end;

procedure TOM_DEVISE.D_PARITEEUROChange;
var St : String ;
begin
     St := GetField('D_PARITEEURO');
     if Valeur(St)>999999.9 then begin
        HShowMessage(msg14, '', '');
        While Valeur(St)>999999.9 do Delete(St,Length(St),1);
        SetField('D_PARITEEURO', St);
     end;
end;

procedure TOM_DEVISE.OnClose;
begin
     if FAvertir then
        ChargeSocieteHalley;
     inherited;
end;


{ TOM_CHANCELL }

procedure TOM_CHANCELL.OnArgument(Arguments: String);
var s, s1, DateD, DateF: string;
begin
  inherited;
  Action := taConsult;
  SurEuro := false;
  s := uppercase(Arguments);
  while s <> '' do begin
    s1 := ReadTokenSt(s);
    if s1 = 'ACTION=CREATION' then
         Action := taCreat
    else if s1 = 'ACTION=MODIFICATION' then
         Action := taModif
    else if s1 = 'SUREURO=TRUE' then
         SurEuro := true;
  end;

  DateD := '''' + DateToStr(VH^.Encours.Deb) + '''';
  DateF := '''' + DateToStr(VH^.Encours.Fin) + '''';
  if DS is TTable then begin
     TTable(DS).Filter := 'H_DATECOURS>=' + DateD +' AND H_DATECOURS<=' + DateF;
     TTable(DS).Filtered:=TRUE;
  end ;

  PDevise   := TPanel(GetControl('PDevise'));
  FListe    := THDBGrid(GetControl('FListe'));
  RGSens    := THRadioGroup(GetControl('RGSens'));
  H_Devise  := THValComboBox(GetControl('H_Devise'));
  TQuotite  := THLabel(GetControl('TQuotite'));
  TValeurQ  := THLabel(GetControl('TValeurQ'));
  TValable  := THLabel(GetControl('TValable'));

  BImprimer := TSpeedButton(GetControl('BImprimer'));
  BPurge    := TSpeedButton(GetControl('BPurge'));

  BImprimer.OnClick := BImprimerClick;
  BPurge.OnClick    := BPurgeClick;

end;

procedure TOM_CHANCELL.OnChangeField(F: TField);
begin
  if F.FieldName = 'H_DEVISE' then
     CodeDeviseChange;
  if (F.FieldName = 'H_TAUXREEL') and (GetField('H_TAUXREEL')<0) then
     SetField('H_TAUXREEL', 0);

  inherited;
end;

procedure TOM_CHANCELL.CodeDeviseChange;
begin
     MonnaieIn := EstMonnaieIn(H_Devise.Value);
     if MonnaieIN then begin
        RGSens.ItemIndex:=1;
        RGSens.Visible := False;
     end
     else begin
        RGSens.Visible := True;
        if H_Devise.Value<>'' then begin
           RGsens.Items[0] := FindEtReplace(msc16, '$', H_Devise.Value, True);
           RGsens.Items[1] := FindEtReplace(msc17, '$', H_Devise.Value, True);
        end;
     end;
end;

procedure TOM_CHANCELL.RGSensClick(Sender: TObject);
begin
     GereSens(True);
end;


procedure TOM_CHANCELL.GereSens(ForceCol: boolean);
begin
     if ForceCol then FListe.SelectedIndex:=3 ;
     FListe.Columns[1].ReadOnly:=(RGSens.ItemIndex=1) ;
     FListe.Columns[2].ReadOnly:=Not (FListe.Columns[1].ReadOnly) ;
     FListe.Columns[1].Visible:=Not FListe.Columns[1].ReadOnly ;
     FListe.Columns[2].Visible:=Not FListe.Columns[2].ReadOnly ;
end;

procedure TOM_CHANCELL.BPurgeClick(Sender: TObject);
var LesDatesASup : TDateTime ;
    Depuis : String ;
begin
     if GetField('H_DATECOURS')<>'' then begin
        Depuis := GetField('H_DATECOURS');
        if PurgeOui(Depuis, H_Devise.Value, LesDatesASup) then
            Purgedates(LesDatesASup);
     end;
end;

procedure TOM_CHANCELL.PurgeDates(DatesSup: TDateTime);
begin
     ExecuteSQl('DELETE FROM CHANCELL WHERE H_DEVISE="'+H_Devise.Value+'" '+
                'AND H_DATECOURS<="'+USDatetime(DatesSup)+'" ');
end;

procedure TOM_CHANCELL.BImprimerClick(Sender: TObject);
begin
     PrintDBGrid (FListe, PDevise, Ecran.Caption, '');
end;

procedure TOM_CHANCELL.OnLoadRecord;
var Quotite : Integer ;
begin
  inherited;
  InitCaption(Ecran, H_Devise.Value, H_Devise.text);
  Quotite := QuotiteChange;
  TQuotite.Visible := Quotite<>1;
  TValeurQ.Visible := Quotite<>1 ;
  if MonnaieIn then
    TValable.Caption := msc10 + '  '+ DateToStr(V_PGI.DateDebutEuro-1)+ '  ' + msc12
  else begin
    if SurEuro then
       TValable.Caption := msc11 + '  ' + DateToStr(V_PGI.DateDebutEuro) + '  ' + msc13
    else
       TValable.Caption := msc10 + '  ' + DateToStr(V_PGI.DateDebutEuro-1)+'  ' + msc12;
  end;
  TValeurQ.Caption := ' ' + IntToStr(Quotite);
  if Action=taConsult then
     BPurge.Enabled := False;
end;


function TOM_CHANCELL.QuotiteChange: Integer;
var Q : TQuery ;
begin
     Result:=1 ;
     Q := OpenSQL('SELECT D_QUOTITE FROM DEVISE WHERE D_DEVISE="'+H_Devise.Value+'"', true);
     if not Q.Eof then
        Result := Q.Fields[0].AsInteger;
     Ferme(Q) ;
end;

Initialization
registerclasses([TOM_DEVISE]);
registerclasses([TOM_CHANCELL]);


end.
