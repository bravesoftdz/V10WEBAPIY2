{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONTROLECBT ()
Mots clefs ... : TOF;BTCONTROLECBT
*****************************************************************}
Unit BTCONTROLECBT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     HTB97,
{$IFNDEF EAGLCLIENT}
		 FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 Maineagl
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB ;

Type

  TCntCbt = (TccNone,tccOk,TccDoubleLiv,TccLivre);

  TControleCbt = class
  private
  	fCodeRet : TcntCbt;
    fTOBD : TOB;
  public
  	property Retour : TCntCbt read fCodeRet write fCodeRet;
    property LesDocs : TOB read fTOBD Write fTOBD;
    procedure Init;
  	constructor create;
    destructor destroy; override;
  end;

  TOF_BTCONTROLECBT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	MRESULT : Tmemo;
    BImprimer : TToolbarButton97;
    TOBLignes : TOB;
    procedure LanceTraitementControle;
    procedure VerifLigneBesoin (TOBL : TOB; Coderetour : TControleCbt);
    procedure TraiteErreur (TOBL : TOB; Resultat : TControleCbt);
    procedure BimprimerClick (Sender : Tobject);
  end ;

procedure VerifBesoinsChantiers;

Implementation
uses factcomm, BTPUtil;

procedure VerifBesoinsChantiers;
begin
//
end;


procedure TOF_BTCONTROLECBT.LanceTraitementControle;
var Chaine : string;
	  QQ : TQuery;
    indice : integer;
    ResControle : TControleCbt;
begin
	TOBLignes.ClearDetail;
	ResControle := TControleCbt.create;
  Chaine := 'Début du contrôle des besoins de chantier à :'+TimeToStr (Now);
	MResult.Clear;
	MResult.lines.Add (Chaine);
	MResult.lines.Add ('--------');
  QQ := OpenSql ('SELECT GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMORDRE,GL_NUMLIGNE,GL_CODEARTICLE,GL_VIVANTE,GL_QTEFACT,GL_QTERESTE,GL_MTRESTE,GL_SOLDERELIQUAT FROM LIGNE WHERE '+
  							 'GL_NATUREPIECEG="CBT" AND GL_QTEFACT <> 0',true);
  if not QQ.eof then
  begin
  	TOBLignes.LoadDetailDB ('LIGNE','','',QQ,false,true);
    ferme (QQ);
    for indice := 0 to TOBLignes.detail.count -1 do
    begin
  		ResControle.init;
      VerifLigneBesoin (TOBLignes.detail[Indice], ResCOntrole);
      TraiteErreur (TOBLignes.detail[Indice],Rescontrole);
    end;
    MResult.lines.Add ('Traitement terminé à '+TimeToStr (Now));
    MResult.lines.Add ('--------');
  end else
  begin
    ferme (QQ);
    MResult.lines.Add ('Rien à traiter');
    MResult.lines.Add ('--------');
  end;
  ResControle.free;
end;

procedure TOF_BTCONTROLECBT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLECBT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLECBT.OnUpdate ;
begin
  Inherited ;
  LanceTraitementControle;
end ;

procedure TOF_BTCONTROLECBT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLECBT.OnArgument (S : String ) ;
begin
  Inherited ;
  MRESULT := Tmemo(GetControl('MRESULT'));
  TOBLignes := TOB.Create ('LES LIGNES',nil,-1);
  BImprimer := TToolbarButton97 (GetControl('Bimprimer'));
  BImprimer.onclick := BimprimerClick;
end ;

procedure TOF_BTCONTROLECBT.OnClose ;
begin
  Inherited ;
  TOBLignes.free;
end ;

procedure TOF_BTCONTROLECBT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLECBT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLECBT.VerifLigneBesoin(TOBL: TOB; CodeRetour : TControleCbt);
var refPiece : string;
		QQ : Tquery;
    TOBPieces : TOB;
    Indice : integer;
    QteLiv : double;
    Ok_ReliquatMt : Boolean;
    MtLiv         : Double;
begin
	CodeRetour.Retour := tccOk;

	TOBPieces := TOB.Create ('LES PIECES',nil,-1);

  RefPiece := EncodeRefPiece (TOBL);

  QQ := OpenSql ('SELECT GL_LIBELLE,GL_QTEFACT,GL_NUMERO,GL_NUMLIGNE FROM LIGNE WHERE GL_NATUREPIECEG="LBT" AND GL_PIECEORIGINE="'+RefPiece+'" AND GL_QTEFACT <> 0',True);

  if not QQ.eof then
  begin
  	TOBPieces.LoadDetailDB ('LIGNE','','',QQ,false);
    ferme (QQ);
    QteLiv  := 0;
    MtLiv   := 0;

    for indice := 0 to TOBPieces.detail.count -1 do
    begin
    	QteLiv := QteLiv + TOBPieces.detail[Indice].GetValue('GL_QTEFACT');
      MtLiv  := MtLiv  + TOBPieces.detail[Indice].GetValue('GL_MONTANTHTDEV');
    end;

    if QteLiv > 0 then
    begin
      if (TOBL.GetValue('GL_QTERESTE') > 0) and (TOBL.GetValue('GL_SOLDERELIQUAT') <> 'X') and (QteLiv > TOBL.GetValue('GL_QTERESTE')) then
      begin
        // cas de la ligne considéré comme non livré dans le besoin avec des reference de livraison
        CodeRetour.Retour := TccLivre;
      end;
      if (TOBL.GetValue('GL_QTERESTE') = 0) and  (QteLiv > TOBL.GetValue('GL_QTEFACT')) then
      begin
        // cas de la ligne traité dont la qteliv > qte initiale
        CodeRetour.Retour := TccLivre;
      end;
      if (TOBL.GetValue('GL_SOLDERELIQUAT') = 'X') and (QteLiv > (TOBL.GetValue('GL_QTEFACT')-TOBL.GetValue('GL_QTERESTE'))) then
      begin
        // cas ou la ligne est considéré comme livré et soldé et dont la quantité de livraison est supérieure a celle considéré comme livré
        CodeRetour.Retour := TccLivre;
      end;
    end;

    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      if MtLiv > 0 then
      begin
        if (TOBL.GetValue('GL_MTRESTE') > 0) and (TOBL.GetValue('GL_SOLDERELIQUAT') <> 'X') and (MtLiv > TOBL.GetValue('GL_MTRESTE')) then
        begin
          // cas de la ligne considéré comme non livré dans le besoin avec des reference de livraison
          CodeRetour.Retour := TccLivre;
        end;
        if (TOBL.GetValue('GL_MTRESTE') = 0) and  (QteLiv > TOBL.GetValue('GL_QTEFACT')) then
        begin
          // cas de la ligne traité dont la qteliv > qte initiale
          CodeRetour.Retour := TccLivre;
        end;
        if (TOBL.GetValue('GL_SOLDERELIQUAT') = 'X') and (MtLiv > (TOBL.GetValue('GL_MONTANTHTDEV')-TOBL.GetValue('GL_MTRESTE'))) then
        begin
          // cas ou la ligne est considéré comme livré et soldé et dont la quantité de livraison est supérieure a celle considéré comme livré
          CodeRetour.Retour := TccLivre;
        end;
      end;
    end;


    if CodeRetour.retour <> TccOk then
    begin
      REPEAT
      	TOBPieces.detail[0].ChangeParent (Coderetour.LesDocs ,-1);
      UNTIL TOBPieces.detail.count = 0 ;
    end;
  end
  else
    ferme (QQ);

  TOBPieces.free;

end;

{ TControleCbt }

constructor TControleCbt.create;
begin
	fTOBD := TOB.Create ('LES DOCUMENTS',nil,-1);
end;

destructor TControleCbt.destroy;
begin
  inherited;
	fTOBD.free;
end;

procedure TControleCbt.Init;
begin
	fTOBD.clearDetail;
  fCodeRet := TccNone;
end;

procedure TOF_BTCONTROLECBT.TraiteErreur(TOBL: TOB; Resultat: TControleCbt);
var DebutMess : String;
		Indice : integer;
begin
	if resultat.Retour = tccOk then exit;
  DebutMess := 'Erreur : Article '+TOBL.GetValue('GL_CODEARTICLE')+' ligne '+IntToStr(TOBL.getValue('GL_NUMLIGNE'))+' dans besoin N° '+InttoStr(TOBL.GetValue('GL_NUMERO'));
  if Resultat.retour = TccDoubleLiv then
  begin
  	DebutMess := DebutMess+' Livré plusieurs fois dans livraisons';
  end else
  begin
  	DebutMess := DebutMess+' sortit dans livraison(s)';
  end;
  for Indice := 0 to resultat.LesDocs.Detail.count -1 do
  begin
    if Indice = 0 then DebutMess := DebutMess + ' '
                  else DebutMess := DebutMess + ',';
    DebutMess := DebutMess+InttoStr(resultat.LesDocs.Detail[Indice].GetValue('GL_NUMERO'));
  end;
  MRESULT.lines.add (DebutMess);
end;

procedure TOF_BTCONTROLECBT.BimprimerClick(Sender: Tobject);
begin
//
end;

Initialization
  registerclasses ( [ TOF_BTCONTROLECBT ] ) ;
end.

