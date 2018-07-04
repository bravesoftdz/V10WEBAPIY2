{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 30/12/2008
Modifié le ... : 19/01/2009
Description .. : Source TOF de la FICHE : BTINFOLIVRAISONS ()
Mots clefs ... : TOF;BTINFOLIVRAISONS
*****************************************************************}
Unit BTINFOLIVRAISONS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     EntGc,
     HMsgBox,
     SaisUtil,
     calcOleGenericBTP,
     UTOF ;

Type
  TOF_BTINFOLIVRAISONS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure AlimenteMemoLiv;
  end ;

function ControlePieceDateModifie (Action : TActionFiche; TOBPiece : TOB) : boolean;
procedure InitLivraisons;
procedure ChangeChantierPiece(Action : TActionFiche; TOBPiece : TOB);
procedure EnregistrePrixLivAModifier (TOBL : TOB);
procedure EnregistreLigneLivAModifier (TOBL : TOB; Qte,PA : double ; WithDetail : boolean=true; Action : string='MODIF');
procedure EnregistreLigneLivASupprimer (TOBL : TOB);
procedure EnregistreDocumentASupprimer (TOBPiece : TOB);
procedure ShowInfoLivraisons;
procedure EnregistreLigneLivAAjouter (TOBL : TOB);

Implementation
var TOBLivraisons : TOB;

procedure AddlesSupLiv (TOBL : TOB);
begin
  TOBL.AddChampSupValeur ('NATUREPIECE','');
  TOBL.AddChampSupValeur ('SOUCHE','');
  TOBL.AddChampSupValeur ('NUMERO',0);
  TOBL.AddChampSupValeur ('INDICE',0);
  TOBL.AddChampSupValeur ('NUMORDRE',0);
  TOBL.AddChampSupValeur ('NUMORDREORIGINE',0);
  TOBL.AddChampSupValeur ('DATEPIECE',StrToDate('01/01/1900'));
  TOBL.AddChampSupValeur ('ARTICLE','');
  TOBL.AddChampSupValeur ('LIBELLE','');
  TOBL.AddChampSupValeur ('AFFAIRE','');
  TOBL.AddChampSupValeur ('QTE',0);
  TOBL.AddChampSupValeur ('PA',0);
  TOBL.AddChampSupValeur ('ACTION','');
end;

procedure InitLivraisons;
begin
  TOBLivraisons.clearDetail;
end;

procedure ChangeChantierPiece(Action : TActionFiche; TOBPiece : TOB);
var ii : Integer;
		TOBL : TOB;
begin
  if (Action = TaModif) and (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),GetPieceAchat (false,false,false,True))>0)  then
  begin
    for II := 0 to TOBPiece.Detail.Count -1 do
    begin
      TOBL := TOBPiece.detail[II];
      if TOBL.GetString('GL_TYPELIGNE')<>'ART' then Continue;
        //PGIInfo (TraduireMemoire ('Attention : Cet article à déjà été livré.#13#10Les consommations ne seront pas mises à jour.#13#10Veuillez mettre à jour les livraisons de chantiers'),caption);
      EnregistreLigneLivAModifier (TOBL,0,0,False,'AFFAIREMODIFIE');
      Break;
    end;
  end;
end;

function ControlePieceDateModifie (Action : TActionFiche; TOBPiece : TOB) : boolean;
var ii : Integer;
		TOBL : TOB;
begin
  result := false;
  if (Action = TaModif) and (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),GetPieceAchat (false,false,false,True))>0)  then
  begin
    for II := 0 to TOBPiece.Detail.Count -1 do
    begin
      TOBL := TOBPiece.detail[II];
      if TOBL.GetString('GL_TYPELIGNE')<>'ART' then Continue;
      if (TOBL.GetValue('BCO_TRAITEVENTE') = 'X' ) then
      begin
        //PGIInfo (TraduireMemoire ('Attention : Cet article à déjà été livré.#13#10Les consommations ne seront pas mises à jour.#13#10Veuillez mettre à jour les livraisons de chantiers'),caption);
        result := true;
				EnregistreLigneLivAModifier (TOBL,0,0,False,'DATEMODIFIE');
      end;
      Break;
    end;
  end;
end;

procedure EnregistreLigneLivASupprimer (TOBL : TOB);
begin
  EnregistreLigneLivAModifier (TOBL,0,0,true,'SUPPR');
end;


procedure EnregistrePrixLivAModifier (TOBL : TOB);
var TOBLL : TOB;
begin
  TOBLL := TOBLivraisons.findFirst(['NUMORDREORIGINE'],[TOBL.getValue('GL_NUMORDRE')],true);
  if TOBLL <> nil then
  begin
    TOBLL.PutValue('PA',TOBL.getValue('GL_PUHTNET'));
  end;
end;

procedure EnregistreLigneLivAAjouter (TOBL : TOB);
var TOBLL : TOB;
begin
  TOBLL := TOB.Create ('UNE LIGNE',TOBLivraisons,-1);
  AddlesSupLiv (TOBLL);
  TOBLL.putValue('NATUREPIECE',TOBL.getValue('GL_NATUREPIECEG'));
  TOBLL.putValue('SOUCHE',TOBL.getValue('GL_SOUCHE'));
  TOBLL.putValue('NUMERO',TOBL.getValue('GL_NUMERO'));
  TOBLL.putValue('INDICE',TOBL.getValue('GL_INDICEG'));
  TOBLL.putValue('NUMORDREORIGINE',TOBL.getValue('GL_NUMORDRE'));
  TOBLL.PutValue('PA',TOBL.getValue('GL_PUHTDEV'));
  TOBLL.PutValue('QTE',TOBL.getValue('GL_QTEFACT'));
  TOBLL.PutValue('ARTICLE',TOBL.getValue('GL_CODEARTICLE'));
  TOBLL.PutValue('LIBELLE',TOBL.getValue('GL_LIBELLE'));
  TOBLL.PutValue('AFFAIRE',TOBL.getValue('GL_AFFAIRE1')+TOBL.getValue('GL_AFFAIRE2')+TOBL.getValue('GL_AFFAIRE3'));
  TOBLL.putvalue ('ACTION','CREATION');
end;

procedure SuppLigneLivCreer (TOBL : TOB);
var TOBLL : TOB;
begin
  TOBLL := TOBLivraisons.findFirst (['NUMORDREORIGINE'],[TOBL.getValue('GL_NUMORDRE')],true);
  if TOBLL <> nil then TOBLL.free;
end;

procedure ModifLigneLivaCreer (TOBL : TOB);
var TOBLL : TOB;
begin
  TOBLL := TOBLivraisons.findFirst (['NUMORDREORIGINE'],[TOBL.getValue('GL_NUMORDRE')],true);
  if TOBLL <> nil then
  begin
    TOBLL.PutValue('PA',TOBL.getValue('GL_PUHTNET'));
    TOBLL.PutValue('QTE',TOBL.getValue('GL_QTEFACT'));
    TOBLL.PutValue('LIBELLE',TOBL.getValue('GL_LIBELLE'));
    TOBLL.PutValue('AFFAIRE',TOBL.getValue('GL_AFFAIRE1')+TOBL.getValue('GL_AFFAIRE2')+TOBL.getValue('GL_AFFAIRE3'));
  end;
end;



procedure EnregistreLigneLivAModifier (TOBL : TOB; Qte,PA : double ; WithDetail : boolean=true; Action : string='MODIF');

  function GetNumLivr (TOBL : TOB; var FromConso : double) : double;
  var QQ : TQuery;
  begin
    result := 0;
    QQ := OpenSql ('SELECT BCO_LIENVENTE,BCO_LIENTRANSFORME FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+IntToStr(TOBL.GetValue('BLP_NUMMOUV'))+
                   ' ORDER BY BCO_NUMMOUV,BCO_INDICE DESC',true,1,'',true);
    if not QQ.eof then
    begin
    	FromConso := QQ.findField('BCO_LIENTRANSFORME').asfloat;
      result := QQ.findField('BCO_LIENVENTE').asfloat;
    end;
    ferme (QQ);
  end;

var TOBLL,TOBM : TOB;
    Select : String;
    QQ,Q1 : TQuery;
    NumLivr,FromConso : double;
begin
	if (Valeur(TOBL.getValue('BLP_NUMMOUV'))=0) then
  begin
  	// cas d'une ligne crée dans la modif de reception ou facture
  	if (Action='MODIF') then
    begin
  	  // cas de l'ajout d'une ligne dans un document existant
    	ModifLigneLivaCreer (TOBL);
    end else if (Action='SUPPR') then
    begin
  	  // cas de l'ajout d'une ligne dans un document existant
    	SuppLigneLivCreer (TOBL);
    end;
    exit;
  end;

  TOBM := TOB.create ('UNE LIV',nil,-1);
  FromConso := 0;
  NumLivr := GetNumLivr (TOBL,FromConso);
  // protection sur ligne non livré
  if NumLivr=0 then
  begin
    TOBM.Free;
  	exit;
  end;
  //
  if (Action='SUPPR') and (Pos(TOBL.getValue('GL_NATUREPIECEG'),GetPieceAchat(false,false,false,true,false))>0) and (FromConso <> 0) then
  begin
    QQ := OpenSql ('SELECT BCO_LIENVENTE FROM CONSOMMATIONS WHERE BCO_NUMMOUV='+floattostr(FromConso)+
                   ' AND BCO_LIENVENTE='+floattostr(numlivr),true,1,'',true);
    if not QQ.eof then
    begin
    	// la on est sur que la piece ayant généré la ligne de livraison n'est pas celle sur laquelle on travaille
      ferme (QQ);
      exit;
    end;
    ferme (QQ);
  end;
  Select := 'SELECT BCO_NATUREPIECEG,BCO_SOUCHE,BCO_NUMERO,BCO_INDICEG,BCO_NUMORDRE FROM CONSOMMATIONS WHERE '+
            'BCO_NUMMOUV='+FloatToStr (NumLivr);
  QQ := OpenSql (Select,true,-1,'',true);
  if not QQ.eof then
  begin
    TOBM.SelectDB ('',QQ);
    if WithDetail then
    begin
      TOBLL := TOBLivraisons.findFirst (['NATUREPIECE','SOUCHE','NUMERO','INDICE','NUMORDRE'],
                                          [TOBM.getValue('BCO_NATUREPIECEG'),
                                          TOBM.getValue('BCO_SOUCHE'),
                                          TOBM.getValue('BCO_NUMERO'),
                                          TOBM.getValue('BCO_INDICEG'),
                                          TOBM.getValue('BCO_NUMORDRE')],true);
    end else
    begin
      TOBLL := TOBLivraisons.findFirst (['NATUREPIECE','SOUCHE','NUMERO','INDICE'],
                                          [TOBM.getValue('BCO_NATUREPIECEG'),
                                          TOBM.getValue('BCO_SOUCHE'),
                                          TOBM.getValue('BCO_NUMERO'),
                                          TOBM.getValue('BCO_INDICEG')],true);
    end;
    if TOBLL = nil then
    begin
      TOBLL := TOB.Create ('UNE LIGNE',TOBLivraisons,-1);
      AddlesSupLiv (TOBLL);
      TOBLL.putValue('NATUREPIECE',TOBM.getValue('BCO_NATUREPIECEG'));
      TOBLL.putValue('SOUCHE',TOBM.getValue('BCO_SOUCHE'));
      TOBLL.putValue('NUMERO',TOBM.getValue('BCO_NUMERO'));
      TOBLL.putValue('INDICE',TOBM.getValue('BCO_INDICEG'));
      TOBLL.putValue('NUMORDREORIGINE',TOBL.getValue('GL_NUMORDRE'));
      if WithDetail then
      begin
        Q1 := OpenSql ('SELECT GL_NUMLIGNE FROM LIGNE WHERE GL_NATUREPIECEG="'+ TOBM.getValue('BCO_NATUREPIECEG')+'" AND '+
                       'GL_SOUCHE="'+ TOBM.getValue('BCO_SOUCHE')+'" AND '+
                       'GL_NUMERO='+IntToStr (TOBM.getValue('BCO_NUMERO'))+' AND '+
                       'GL_INDICEG='+IntToStr(TOBM.getValue('BCO_INDICEG'))+' AND '+
                       'GL_NUMORDRE='+IntToStr(TOBM.getValue('BCO_NUMORDRE')),true,-1,'',true);
        if not Q1.eof then
        begin
          TOBLL.putValue('NUMORDRE',Q1.Findfield('GL_NUMLIGNE').AsInteger);
        end;
        ferme(Q1);
      end;
    end;
    TOBLL.AddChampSupValeur ('ACTION',Action);
    if QTe <> 0 then
    begin
      TOBLL.PutValue('QTE',Qte);
    end;
    if PA <> 0 then
    begin
      TOBLL.PutValue('PA',PA);
    end;
    if Action='DATEMODIFIE' then
    begin
      TOBLL.PutValue('DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
    end else if Action='AFFAIREMODIFIE' then
    begin
      TOBLL.PutValue('AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
    end;
  end;
  ferme (QQ);
  TOBM.Free;
end;

procedure EnregistreDocumentASupprimer (TOBPiece : TOB);
var Indice : integer;
    TOBLL,TOBL : TOB;
    current,last : string;
begin
  // phase 1 --  raz des numero de lignes déjà présentes
  for Indice := 0 to TOBLivraisons.detail.count -1 do
  begin
    TOBLL := TOBLivraisons.detail[Indice];
    if TOBLL.getValue('NUMORDRE')<>0 then TOBLL.putValue('NUMORDRE',0);
  end;
  // phase 2 --  Tri
  TOBLivraisons.detail.Sort('NATUREPIECE;SOUCHE;NUMERO;INDICE');
  // phase 3 --  Dedoublonnage
  indice := 0;
  current := '';
  Last := '';
  if TOBLivraisons.detail.count > 0 then
  begin
    repeat
      TOBLL := TOBLivraisons.detail[Indice];
      Current := TOBLL.GetValue('NATUREPIECE')+';'+TOBLL.GetValue('SOUCHE')+';'+IntToStr(TOBLL.GetValue('NUMERO'))+
                 IntToStr(TOBLL.GetValue('INDICE'));
      if Current = Last then
      begin
        TOBLL.free
      end else
      begin
        Last := TOBLL.GetValue('NATUREPIECE')+';'+TOBLL.GetValue('SOUCHE')+';'+IntToStr(TOBLL.GetValue('NUMERO'))+
                IntToStr(TOBLL.GetValue('INDICE'));
        Inc(Indice);
      end;
    until Indice >= TOBLivraisons.detail.Count;
  end;
  // phase 4 -- parcours de la piece pour memorisation des documents a supprimer
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL.getValue('GL_TYPELIGNE')<>'ART' then continue;
    EnregistreLigneLivAModifier (TOBL,0,0,False,'SUPPR');
  end;
end;

procedure ShowInfoLivraisons;
begin
  if TOBLivraisons.detail.Count = 0 then exit;
  AGLLanceFiche ('BTP','BTINFOLIVRAISONS','','','');
end;

//
procedure TOF_BTINFOLIVRAISONS.AlimenteMemoLiv;

    procedure AddLigne(MM: Tmemo;TOBL : TOB);
    var St,St1 : String;
    begin

      if TOBL.getValue('NUMORDRE') <> 0 then
      begin
        St1 := 'la ligne n° '+IntToStr(TOBL.getValue('NUMORDRE'))+' de la livraison n° '+InttOStr(TOBL.getValue('NUMERO'));
      end else
      begin
        St1 := 'livraison n° '+InttOStr(TOBL.getValue('NUMERO'));
      end;

      if TOBL.GetValue('ACTION')='DATEMODIFIE' then
      begin
        St := 'Modifier '+St1+ ' - Date à passer au '+DateToStr(TOBL.Getvalue('DATEPIECE'));
      end else if TOBL.GetValue('ACTION')='SUPPR' then
      begin
        St := 'Supprimer '+St1;
      end else if TOBL.GetValue('ACTION')='MODIF' then
      begin
        St := 'Modifier '+St1;
        if TOBL.getValue('QTE') <> 0 then
        begin
          St := St + ' - quantité à passer à '+StrF00(TOBL.getValue('QTE'), V_PGI.okdecQ);
        end;
        if TOBL.getValue('PA') <> 0 then
        begin
          St := St + ' - prix d''achat à passer à '+StrF00(TOBL.getValue('PA'), V_PGI.okdecP);
        end;
      end else if TOBL.GetValue('ACTION')='AFFAIREMODIFIE' then
      begin
        St := 'Modifier la livraison '+InttOStr(TOBL.getValue('NUMERO'))+' --> Chantier à passer sur '+BTPCodeAffaireAffiche (TOBL.GetString('AFFAIRE'),'');
      end else
      begin
      	st := 'Créer une livraison pour l''article "'+TOBL.getValue('ARTICLE')+' '+TOBL.getValue('LIBELLE')+'" pour l''affaire '+TOBL.getValue('AFFAIRE');
        MM.lines.add(St);
      	st := 'd''une quantité de '+StrF00(TOBL.getValue('QTE'), V_PGI.okdecQ)+' au prix de '+StrF00(TOBL.getValue('PA'), V_PGI.okdecP);
      end;
      MM.lines.add(St);
    end;

var MM : Tmemo;
    Indice : integer;
    TOBL : TOB;
begin
  MM := Tmemo(GetControl('MEMOLIV'));
  MM.Lines.Clear;
  TOBLivraisons.detail.Sort('NATUREPIECE;SOUCHE;NUMERO;INDICE;NUMORDRE');
  for Indice := 0 to TOBLivraisons.detail.count -1 do
  begin
    TOBL := TOBLivraisons.detail[Indice];
    AddLigne(MM,TOBL);
  end;
end;

procedure TOF_BTINFOLIVRAISONS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnArgument (S : String ) ;
begin
  Inherited ;
  AlimenteMemoLiv;
end ;

procedure TOF_BTINFOLIVRAISONS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTINFOLIVRAISONS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTINFOLIVRAISONS ] ) ;
  TOBLivraisons := TOB.Create ('LES LIGNES DE LIV',nil,-1);

FINALIZATION
  TOBLivraisons.free;

end.
