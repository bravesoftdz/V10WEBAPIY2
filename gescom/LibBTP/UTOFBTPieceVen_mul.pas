{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/08/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECEVEN_MUL ()
Mots clefs ... : TOF;BTPIECEVEN_MUL
*****************************************************************}
Unit UTOFBTPieceVen_mul ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UtilsRapport,
     uTofAfBaseCodeAffaire,
     UTOF ;

Type
  TOF_BTPIECEVEN_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

private

  BOuvrir      : TToolBarButton97;
  RapportGen   : TGestionRapport;
  //
  procedure Validation(Sender: Tobject);

  end ;

Implementation

uses UtilsOuvragesPlat,
     uEntCommun;

procedure TOF_BTPIECEVEN_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnArgument (S : String ) ;
begin
  Inherited ;

  BOuvrir := TToolbarButton97(GetControl('BOUVRIR'));
  BOuvrir.OnClick := Validation;
  //
  RapportGen   := TGestionRapport.Create(TFMul(Ecran));
  with RapportGen do
  begin
    Titre   := 'Génération des ouvrages à plat';
    Affiche := True;
    Close   := True;
    Sauve   := False;
    Print   := False;
    InitRapport;
    Visible := false;
    PosLeft := Round(TFMul(Ecran).Width / 1.5);
    PosTop  := Round(TFMul(Ecran).Top);
  end;

  //Gestion du code affaire....


end ;

procedure TOF_BTPIECEVEN_MUL.OnClose ;
begin
  RapportGen.free;
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECEVEN_MUL.Validation(Sender : Tobject);
var F       : TFMul ;
    QQ      : TQuery;
    cledoc  : r_cledoc;
    XX      : TGenOuvPlat;
    i       : Integer;
begin

  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIError(F.Caption,'Aucun élément sélectionné') ;
    exit;
  end;

  RapportGen.InitRapport;

  XX := TGenOuvPlat.create;
  TRY
    if F.FListe.AllSelected then
    begin
      QQ:=F.Q;
      QQ.First;
      while Not QQ.EOF do
      Begin
        XX := TGenOuvPlat.create;
        //
        cledoc.NaturePiece  := QQ.FindField('GP_NATUREPIECEG').AsString;
        cledoc.DatePiece    := QQ.FindField('GP_DATEPIECE').AsDateTime;
        Cledoc.Souche       := QQ.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece  := QQ.FindField('GP_NUMERO').AsInteger;
        cledoc.Indice       := QQ.FindField('GP_INDICEG').AsInteger;
        //
        XX.TraitePiece(cledoc);
        //
        if (F.FListe.NbSelected=1) then
          PGIInfo(XX.LibRetour)
        else
          RapportGen.SauveLigMemo(XX.LibRetour);
        //
        QQ.Next;
      end;
      F.FListe.AllSelected:=False;
    end
    else
    begin
      for i:=0 to F.FListe.NbSelected-1 do
      begin
        F.FListe.GotoLeBookmark(i);
        //
        cledoc.NaturePiece  := F.FListe.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        cledoc.DatePiece    := F.FListe.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
        Cledoc.Souche       := F.FListe.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece  := F.FListe.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.Indice       := F.FListe.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        //
        XX.TraitePiece(cledoc);
        if (F.FListe.NbSelected=1) then
          PGIInfo(XX.LibRetour)
        else
          RapportGen.SauveLigMemo(XX.LibRetour);
      end;
      F.FListe.ClearSelected;
    end;
  FINALLY
    if RapportGen.Memo.Lines.count > 0 then
    begin
      RapportGen.AfficheRapport;    //affichage du rapport d'intégration
      RapportGen.Affiche := True;
    end;
    XX.Free;
  END;

end;

procedure TOF_BTPIECEVEN_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ; Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ; Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
// tiers non passé sur les écrans d'achats car non maj / tiers de l'affaire
if Pos('VEN',Ecran.Name)= 0 then Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end;

Initialization
  registerclasses ( [ TOF_BTPIECEVEN_MUL ] ) ;
end.

