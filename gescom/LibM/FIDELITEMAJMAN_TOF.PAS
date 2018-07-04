{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 10/07/2003
Modifié le ... : 10/07/2003
Description .. : Source TOF de la FICHE : FIDELITEMAJMAN ()
Mots clefs ... : TOF;FIDELITEMAJMAN
*****************************************************************}
Unit FIDELITEMAJMAN_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
     MaineAgl,
  {$ELSE}
     db, dbtables, Fe_Main,
  {$ENDIF}
  {$IFDEF FOS5}
     FIDELITEINFO_TOF, ENTGC,
  {$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
     FOUtil, UTOB, UFidelite, Vierge, M3FP ;

function Lance_MAJ_Fidelite(NumCarteInt, Vendeur, Etab : String; AncienCumul : Double) : Double;

Type
  TOF_FIDELITEMAJMAN = Class (TOF)
      procedure OnNew                    ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnClose                  ; override ;
    Private
      Etab, NumCarteInterne, Prog, TypeCumul, Representant : string;
      OldCumul : Double;
      OkFerme : Boolean;
      procedure GFI_VALEUR_OnExit(Sender: TObject);
      procedure NEWCUMUL_OnExit(Sender: TObject);
  end ;

Implementation

procedure TOF_FIDELITEMAJMAN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_FIDELITEMAJMAN.OnUpdate ;
var TOBFID : TOB;
    NewValeur : Double;
begin
  Inherited ;
  OkFerme := True;
  NewValeur := Valeur(GetControlText('GFI_VALEUR'));

  if NewValeur <> 0.0 then
  begin
    if GetControlText('GFI_COMMENTAIRE') = '' then
    begin
      PGIError('Le commentaire est obligatoire.');
      SetFocusControl('GFI_COMMENTAIRE');
      OkFerme := False; Exit;
    end;
    TOBFID := TOB.Create('FIDELITELIG', Nil, -1);
    With TOBFID do
    begin
      PutValue('GFI_CAISSE', FOCaisseCourante);
      PutValue('GFI_COMMENTAIRE', GetControlText('GFI_COMMENTAIRE'));
      PutValue('GFI_CREATEUR', V_PGI.User);
      PutValue('GFI_DATECREATION', NowH);
      PutValue('GFI_DATEINTEGR', iDate1900);
      PutValue('GFI_DATEMODIF', NowH);
      PutValue('GFI_ETABLISSEMENT', Etab);
      PutValue('GFI_LIGNE', RecupChronoParamSoc('SO_CHRONOFIDELLIG'));
      PutValue('GFI_NATUREPIECEG', '');
      PutValue('GFI_NUMCARTEINT', NumCarteInterne);
      PutValue('GFI_NUMERO', 0);
      PutValue('GFI_NUMORDRE', 0);
      PutValue('GFI_PROGRAMME', Prog);
      PutValue('GFI_REPRESENTANT', Representant);
      PutValue('GFI_SOUCHE', '');
      PutValue('GFI_TYPECUMULFID', TypeCumul);
      PutValue('GFI_TYPELIGNEFIDEL', '888'); //Saisie manuelle
      PutValue('GFI_UTILISATEUR', V_PGI.User);
      PutValue('GFI_VALEUR', NewValeur);
      if InsertDB(Nil) then TFVierge(Ecran).Retour := FloatToStr(NewValeur);
      Free;
    end;
  end;

end ;

procedure TOF_FIDELITEMAJMAN.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_FIDELITEMAJMAN.OnArgument (S : String ) ;
var Critere, StChamp, StValeur : String;
    x : Integer;
    Q_FIDELITEENT, Q_PARFIDELITE : TQuery;
begin
  Inherited ;

  THEdit(GetControl('NEWCUMUL')).OnExit := NEWCUMUL_OnExit;
  THEdit(GetControl('GFI_VALEUR')).OnExit := GFI_VALEUR_OnExit;

  TFVierge(Ecran).Retour := '0';
  OkFerme := True;

  Repeat
    Critere := ReadTokenSt(S) ;
    if Critere<>'' then
    begin
      x := pos('=',Critere);
      if x<>0 then
      begin
        StChamp := Copy(Critere, 1, x-1);
        StValeur := copy(Critere,x+1,length(Critere));
        if StChamp = 'NUMEROCARTEINTERNE' then NumCarteInterne := StValeur
        else if StChamp = 'VENDEUR' then Representant := StValeur
        else if StChamp = 'OLDCUMUL' then
        begin
          SetControlText( 'OLDCUMUL', StValeur );
          SetControlText( 'NEWCUMUL', StValeur );
          OldCumul := StrToFloat(StValeur) ;
        end;
      end;
    end;
  until Critere='';

  Q_FIDELITEENT := OpenSql('SELECT GFE_ETABLISSEMENT, GFE_PROGRAMME ' +
  'FROM FIDELITEENT WHERE GFE_NUMCARTEINT="' + NumCarteInterne + '"', True);
  if Not Q_FIDELITEENT.EOF then
  begin
    With Q_FIDELITEENT do
    begin
      {$IFDEF FOS5}
      Etab := VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT');
      {$ELSE}
      Etab := FindField('GFE_ETABLISSEMENT').AsString;
      {$ENDIF}
      Prog := FindField('GFE_PROGRAMME').AsString;
    end;
  end;
  Ferme(Q_FIDELITEENT);

  Q_PARFIDELITE := OpenSql('SELECT GFO_TYPECUMULFID FROM PARFIDELITE WHERE GFO_CODEFIDELITE="' + Prog + '"', True);
  if Not Q_PARFIDELITE.EOF then TypeCumul := Q_PARFIDELITE.FindField('GFO_TYPECUMULFID').AsString;
  Ferme(Q_PARFIDELITE);

  SetControlText( 'GFI_VALEUR', '0' );
  SetControlText( 'GFI_COMMENTAIRE', '' );
  SetControlText( 'GFI_DATECREATION', DateToStr(NowH) );
  SetControlText( 'GFI_REPRESENTANT', Representant );
  SetControlText( 'GFI_CREATEUR', V_PGI.User );

end ;

procedure TOF_FIDELITEMAJMAN.OnClose ;
begin
  Inherited ;
  if OkFerme then LastError := 0
  else
  begin
    LastError := -1;
    OkFerme := True;
  end;
end ;

// Le résultat est la valeur ajouté à la fidélité
function Lance_MAJ_Fidelite(NumCarteInt, Vendeur, Etab : String; AncienCumul : Double) : Double;
var Arguments : String;
begin
  if Etab <> '' then
  begin //Charge l'ancien cumul de la carte
    With Fidelite.Create do
    begin
      LoadProgramme('#CLIENTDEFAUT#', Etab);
      if ChargeCarte('', '#CLIENTDEFAUT#', Etab, NumCarteInt, '', False) then
        AncienCumul := CumulFidelite;
      Free;
    end;
  end;
  Arguments := 'NUMEROCARTEINTERNE=' + NumCarteInt +
  ';VENDEUR=' + Vendeur +
  ';OLDCUMUL=' + FloatToStr(AncienCumul);
  Result := StrToFloat( AGLLanceFiche('MBO','FIDELITEMAJMAN','','', Arguments) );
end;

procedure TOF_FIDELITEMAJMAN.GFI_VALEUR_OnExit(Sender: TObject);
begin
  SetControlText('NEWCUMUL', FloatToStr( OldCumul + Valeur(GetControlText('GFI_VALEUR')) ) );
end;

procedure TOF_FIDELITEMAJMAN.NEWCUMUL_OnExit(Sender: TObject);
begin
  SetControlText('GFI_VALEUR', FloatToStr( Valeur(GetControlText('NEWCUMUL')) - OldCumul ) );
end;

///////////////////////////////////////////////////////////////////////////////

procedure AGLLance_MAJ_Fidelite(parms: array of variant; nb: integer);
var DbRetour : Double;
{$IFDEF FOS5}
  F: TForm;
  TOTOF: TOF;
{$ENDIF}
begin
  {$IFDEF FOS5}
  if Not FOJaiLeDroit(87) then Exit;
  {$ENDIF}

  DbRetour := Lance_MAJ_Fidelite(Parms[1], Parms[2], Parms[3], StrToFloat(Parms[4]));
  if DbRetour = 0.0 then Exit;
  {$IFDEF FOS5}
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTOF else exit;
  if (TOTOF is TOF_FIDELITEINFO) then
  begin
    With TOTOF do
    begin
      SetControlText('VALRETOUR', FloatToStr(DbRetour) );
      SetControlText('TCUMULOLD', FloatToStr(StrToFloat(GetControlText('TCUMULOLD')) + DbRetour) );
      SetControlText('TCUMULNEW', FloatToStr(StrToFloat(GetControlText('TCUMULNEW')) + DbRetour) );
      SetControlChecked('ISMODIFCUMUL',True);
    end;
  end else exit;
  {$ENDIF}
end;

Initialization
  registerclasses ( [ TOF_FIDELITEMAJMAN ] ) ;
  RegisterAglProc('Lance_MAJ_Fidelite', True, 4, AGLLance_MAJ_Fidelite);
end.
