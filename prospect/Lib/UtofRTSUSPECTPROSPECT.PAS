{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/09/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTSUSPECTPROSPECT ()
Mots clefs ... : TOF;RTSUSPECTPROSPECT
*****************************************************************}
Unit UtofRTSUSPECTPROSPECT ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fe_Main,
{$else}
     eMul,
     Maineagl,
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97,
     KpmgUtil,
     Ed_tools,
     Paramsoc,
     utofcomm,
     utilGC,
     ent1,
     UtilRT,
     UtilsRapport,
     UTOF ;

Type
  TOF_RTSUSPECTPROSPECT = Class (tTOFComm)
  	private
    	OneRapport  : TGestionRapport;
    Public

    bOuvrir       : TToolbarButton97;
    bZoom         : TToolbarButton97;
    Fliste        : THGrid;
    GereEnseigne  : Boolean;

    procedure bOuvrirOnClick ( Sender : Tobject );
    procedure bZoomOnClick ( Sender : Tobject );

    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_RTSUSPECTPROSPECT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

  uses
    Variants,UFonctionsCBP
    ;

Function RTLanceFiche_RTSUSPECTPROSPECT(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
  if (not ExJaiLeDroitConcept(TConcept(GcTransfertSusPro), False)) then
    PGIInfo('Vous n''avez pas les droits de transfert de suspect en prospect', 'Transfert par lot de suspects en prospects')
  else
    AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTSUSPECTPROSPECT.OnArgument (Arguments : String ) ;
var
  StrWhere              : String;
  StrMsg                : String;
  bBloque               : boolean;
  BAvert                : Boolean;

begin
  Inherited ;
  StrWhere              := '';
  fTableName            := 'SUSPECTS' ;
  fMulDeTraitement      := True ;
  OneRapport := TGestionRapport.Create(ecran);
  OneRapport.Titre := 'Intégration de suspects';
  OneRapport.Affiche := True;
  OneRapport.Close := True;



  bOuvrir               := TToolbarButton97(GetControl('BOUVRIR'));
  bOuvrir.OnClick       := bOuvrirOnClick;

  //FQ10464 changement des messages d'informations
  StrMsg                := JaiLeDroitCiblage(bBloque, BAvert);
  if StrMsg <> '' then
  begin
    if bBloque then bouvrir.visible := False;
    PGIInfo(StrMsg, TFMul(Ecran).Caption);
  end;

  bZoom                 := TToolbarButton97(GetControl('BZOOM'));
  if bZoom <> nil then
    bZoom.OnClick       := bZoomOnClick;


  Fliste                := THGrid(GetControl('FLISTE'));
  Fliste.OnDblClick     := bZoomOnClick;

  SetControlText('RSU_FERME', '-');
  SetControlText('XX_WHERE', StrWhere);
  RTMajChampsLibresSuspect(TFmul(Ecran));

  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('RSU_REPRESENTANT', False);
    SetControlVisible ('TRSU_REPRESENTANT', False);
  end;
  
end ;

procedure TOF_RTSUSPECTPROSPECT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnClose ;
begin
	OneRapport.Free;
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTSUSPECTPROSPECT.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_RTSUSPECTPROSPECT.bOuvrirOnClick(Sender: Tobject);
var
   F : TFMul;
   Lgfiche : integer;
   LeTiers              : String;
   LeSuspect            : String;
   Qry                  : TQuery;
begin

  F := TFMul(Ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
    messagealerte('Aucun élément sélectionné');
    exit;
  end;
  OneRapport.InitRapport;
  TRY
    {
    //GetParamSocSecur('SO_BOURREAUX','X',True);
    Cbour := GetParamSocSecur('SO_BOURREAUX','X',True);
    Cbour := UpperCase_(Cbour);
    if not (Cbour[1] in ['A'..'Z']) then
    begin
    PGIInfo('Attention !#10#13Le caractère de bourrage du compte auxiliaire est numérique.#10#13'+
    'Le transfert de suspect en prospect est impossible.', 'Avertissement');
    exit;
    end;


    if PGIAsk('Attention !#10#13La création par lot de fiches prospects implique que :#10#13'+
    '- seuls les champs qui ne sont pas concernés par vos restrictions fiches sont mis à jour en série,#10#13'+
    '- les contrôles effectués à la création manuelle d''un prospect (champs obligatoires, tablettes hiérarchiques...) ne sont pas réalisés.#10#13'+
    'Voulez-vous continuez le traitement ?', 'Avertissement') = mrNo then exit;

    //FQ 10479
    GereEnseigne := GetParamSocSecur('SO_GCENSEIGNETAB', False, True);
    }
    if F.Fliste.AllSelected then
    begin
      InitMoveProgressForm(nil,'Traitement en cours...','',F.Q.recordcount,True,True);
      TRY
        F.Q.First;
        While not F.Q.Eof do
        begin
          LeSuspect:= F.Q.FindField('RSU_SUSPECT').AsString;
          MoveCurProgressForm('Transfert suspect : ' + F.Q.FindField('RSU_LIBELLE').AsString);
          LeTiers:= SuspectVersProspect(LeSuspect, True, GereEnseigne,True,OneRapport);    //fonction de passage suspect vers prospect sans créer dans l'annuaire
          if Letiers <> '' then
          begin
            Qry := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + LeTiers + '"', True);
            ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_CONTACT="'+Qry.FindField('T_AUXILIAIRE').AsString+
            '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+LeSuspect+'" AND RVB_NUMEROCONTACT="1"');

            ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_TIERS="'+LeTiers+
            '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+LeSuspect+'" AND RVB_NUMEROCONTACT="0"');
            Ferme(Qry);
          end;
          F.Q.Next;
        end;
      FINALLY
      	FiniMoveProgressForm;
      end;

    end else
    begin
      InitMoveProgressForm(nil,'Traitement en cours...','',F.FListe.nbSelected,True,True);
      TRY
        for lgfiche := 0 to F.Fliste.NbSelected - 1 do
        begin
          MoveCurProgressForm('Transfert suspect : ' + F.Q.FindField('RSU_LIBELLE').AsString);
          F.Fliste.GotoLeBookmark(lgfiche);
          {$IFDEF EAGLCLIENT}
          F.Q.TQ.Seek(F.FListe.row -1);
          {$ENDIF}
          LeSuspect := F.Q.FindField('RSU_SUSPECT').AsString;

          LeTiers := SuspectVersProspect(LeSuspect, True, GereEnseigne,True,OneRapport);    //fonction de passage suspect vers prospect sans créer dans l'annuaire
          if LeTiers <> '' then
          begin
            Qry               := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + LeTiers + '"', True);
            ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_CONTACT="'+Qry.FindField('T_AUXILIAIRE').AsString+
                        '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+LeSuspect+'" AND RVB_NUMEROCONTACT="1"');

            ExecuteSQL ('UPDATE CIBLAGEELEMENT SET RVB_TIERS="'+LeTiers+
                        '", RVB_SUSPECT="XXX" WHERE RVB_SUSPECT="'+LeSuspect+'" AND RVB_NUMEROCONTACT="0"');
            Ferme(Qry);
          end;

        end;
      FINALLY
      	FiniMoveProgressForm;
      end;
    end;
  finally
    OneRapport.AfficheRapport;
    F.BChercheClick(Ecran);
    F.FListe.ClearSelected;
  end;
end;

procedure TOF_RTSUSPECTPROSPECT.bZoomOnClick(Sender: Tobject);
begin
  if VarIsNull(GetField('RSU_SUSPECT')) then
    exit;

  AGLLanceFiche('RT','RTSUSPECTS','',Getfield('RSU_SUSPECT'),'MONOFICHE;ACTION=MODIFICATION');
  RefreshDB;



end;

Initialization
  registerclasses ( [ TOF_RTSUSPECTPROSPECT ] ) ;
end.
