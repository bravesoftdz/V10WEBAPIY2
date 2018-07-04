unit UtofGridFournisseur;

interface
uses  sysutils,controls,Classes,HCtrls,UTof,Utob,Vierge,forms,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fe_Main,
      Db{$IFNDEF DBXPRESS} ,
      dbTables {$ELSE} ,
      uDbxDataSet {$ENDIF},
{$ENDIF}
{$IFDEF BTP}
      ParamSoc,
      UtilTarif,
      HEnt1,
      HTB97,
{$ENDIF}
      Ent1,
      HMsgBox,
      M3FP;

Type
    TOF_GridFournisseur = Class(TOF)
    private
        Rang        : string ;
        TypAff      : string ;
        Article     : string;
        LesColonnes : string;
        BZoom       : TToolBarButton97;
        Procedure OnLoad; override ;
        Procedure OnArgument(Argument:string) ; override ;
{$IFDEF BTP}
        procedure BTPRefreshLaLigne ;
        procedure RefreshData(LaLigne: integer);
{$ENDIF}
        function Insert : variant;
        procedure BZOOM_OnClick(Sender: TObject);
    END;

var LibelleArticle : string ;

implementation
uses UFonctionsCBP;

{$IFDEF BTP}
procedure TOF_GridFournisseur.RefreshData (LaLigne : integer);
var ThisTOb         : TOB;
    TOBART          : TOB;
    TOBTIers        : TOB;
    TOBTarif        : TOB;
    TOBCataFouPrinc : TOB;
    TOBCatalog      : TOB;
    CodArtic        : string;
    MTPAF           : Double;
    PrixPourQte     : double;
    QQ              : TQuery;
begin
  TOBART := TOB.Create ('ARTICLE',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TOBTArif := TOB.Create ('TARIF',nil,-1);
  TOBCataFouPrinc := TOB.Create ('CATALOGU',nil,-1);
  TOBCatalog := TOB.Create ('CATALOGU',nil,-1);

  ThisTOB := LaTOB.Detail[LaLigne -1];
  CodArtic := ThisTOB.GetValue('GCA_ARTICLE');
  TRY
  	QQ := OpenSql ('SELECT * FROM ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE GA_ARTICLE="'+CodArtic+'"',True);
    TOBART.SelectDb('',QQ);
    ferme (QQ);
    //
    TOBTIers.PutValue ('T_AUXILIAIRE',ThisTOB.GetValue('GCA_TIERS') );
    TOBTIERS.loaddb();
    //
    TOBCatalog.PutValue ('GCA_REFERENCE',ThisTOB.GetValue('GCA_REFERENCE'));
    TOBCatalog.PutValue ('GCA_TIERS',ThisTOB.getValue('GCA_TIERS'));
    TOBCatalog.LoadDb();

    if TOBART.GetValue('GA_FOURNPRINC') <> '' then
    begin
      QQ := OpenSql ('SELECT * FROM CATALOGU WHERE GCA_TIERS="'+TOBART.GEtVAlue('GA_FOURNPRINC')+'" AND GCA_ARTICLE="'+CodArtic+'"',True);
      if not QQ.eof Then TOBCataFouPrinc.SelectDB ('',QQ);
      ferme (QQ);
    end;
    if LaTob.getValue ('MODE') then
    begin
      GetTarifGlobal (CodArtic,TOBArt.getValue('GA_TARIFARTICLE'),TOBArt.getValue('GA2_SOUSFAMTARART'),'ACH',TOBArt,TOBTiers,TOBTarif,true);
      if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
      begin
        MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
      end else
      begin
        prixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;
        if (TOBCatalog.GetValue('GCA_PRIXVENTE') = 0) and
           (TOBCatalog.GetValue('GCA_PRIXBASE') = 0) then
        begin
          if TOBCataFouPrinc.GetValue('GCA_PRIXBASE') = 0 then
          begin
            MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXVENTE')/prixPourQte;
          end else
          begin
            MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXBASE')/prixPourQte;
          end;
        end else
        begin
          if TOBCatalog.GetValue('GCA_PRIXBASE') = 0 then
          begin
           MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
          end else
          begin
           MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQTe;
          end;
        end;
      end;
      ThisTOB.PutValue('GCA_PRIXBASE',MTPAF);
      ThisTOB.PutValue('GF_PRIXF',MTPAF);
      ThisTOB.PutValue('GF_REMISE',TOBTarif.getValue('GF_CALCULREMISE'));
      MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecV );
      ThisTOB.PutValue('GF_TARIF',MTPAF);
    end else
    begin
      ThisTOB.PutValue('GCA_PRIXBASE',TOBCatalog.GetValue('GCA_PRIXBASE'));
    end;
    ThisTOB.PutValue('GCA_DELAILIVRAISON',TOBCatalog.GetValue('GCA_DELAILIVRAISON'));
    ThisTob.PutLigneGrid (ThGrid(GetControl('GRIDFOURN')),LaLigne,false,false,LesColonnes);
  FINALLY
    FreeAndNil (TOBART);
    FreeAndNil (TOBTiers);
    FreeAndNil (TOBTarif);
    FreeAndNil (TOBCataFouPrinc);
    FreeAndNil (TOBCatalog);
  END;
end;

procedure TOF_GridFournisseur.BTPRefreshLaLigne;
var LaLigne : integer;
begin
  LaLigne := THGrid(GetControl('GRIDFOURN')).row;
  RefreshData (LaLigne);
//
end;
{$ENDIF}

function TOF_GridFournisseur.Insert : variant;
var stLequel : string;
begin
  stlequel := LibelleArticle;
{$IFDEF BTP}
  if Article <> '' then
  begin
    Stlequel := StLequel + ';ARTICLE='+Article;
  end;
{$ENDIF}
  result:=AGLLanceFiche('GC','GCCATALOGU_SAISI3','',Article ,'ACTION=CREATION;LIB='+stLequel) ;
end;

procedure TOF_GridFournisseur.BZOOM_OnClick(Sender : TObject);
Var Action    : String;
    Fourn     : String;
    Art       : String;
    LaLigne   : Integer;
    TobLigCat : TOB;
begin

  LaLigne   := THGrid(GetControl('GRIDFOURN')).row;
  TobLigCat := LaTOB.Detail[LaLigne -1];

  if ExJaiLeDroitConcept (TConcept(gcCatalogue),False) then
    Action :='ACTION=MODIFICATION'
  else
    Action :='ACTION=CONSULTATION';

  Fourn := TobLigCat.GetString('GCA_TIERS');
  Art   := TobLigCat.GetString('GCA_REFERENCE');

  AGLLanceFiche('GC','GCCATALOGU_SAISI3','',Art+';'+Fourn,Action);

end;

Procedure TOF_GridFournisseur.OnArgument(Argument:string) ;
var critere,ChampMul,ValMul : string ;
    x : integer ;
{$IFDEF BTP}
    GridFourn : THGrid;
    FF : string;
    i_ind1 : integer;
{$ENDIF}
begin
inherited ;

  LesColonnes := '';

  if ecran.name = 'BTGRIDFOURNISSEUR' then
     LesColonnes := 'GCA_TIERS;GCA_REFERENCE;T_LIBELLE;GCA_PRIXBASE;GF_PRIXF;GF_REMISE;GF_TARIF;GCA_DELAILIVRAISON';

  Repeat
    Critere:=Trim(ReadTokenSt(Argument)) ;
    if Critere<>'' then
      begin
      x:=pos('=',Critere) ;
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1) ;
         ValMul:=copy(Critere,x+1,length(Critere)) ;
         if ChampMul='FOURN' then Rang:=uppercase(ValMul) ;
         if ChampMul='LIBART' then LibelleArticle:=ValMul ;
         if ChampMul='ART' then Article:=ValMul ;
         if ChampMul='TYPAFF' then TypAff:=ValMul ;
         end ;
      end ;

  {$IFDEF BTP}
  if ecran.name = 'BTGRIDFOURNISSEUR' then
     begin
     FF := '';
     for i_ind1 := 1 to GetParamSoc('SO_DECPRIX') do FF := FF + '0';
     GridFourn := THGrid(GetControl('GRIDFOURN'));
     GridFourn.FColAligns [GridFourn.colcount-1] := taRightJustify; // delai
     GridFourn.FColAligns [GridFourn.colcount-2] := taRightJustify; // PA NEt
     GridFourn.FColFormats [GridFourn.colcount-2] := '#,##0.' + FF;
     GridFourn.FColAligns [GridFourn.colcount-4] := taRightJustify; // prix de base
     GridFourn.FColFormats [GridFourn.colcount-4] := '#,##0.' + FF;

     GridFourn.FColAligns [GridFourn.colcount-5] := taRightJustify; // tarif de base
     GridFourn.FColFormats [GridFourn.colcount-5] := '#,##0.' + FF;
     if (not GetParamSoc ('SO_BTRECHTARIFFOU')) then
        BEGIN
        GridFourn.ColWidths [GridFourn.colcount-1] := 0;
        GridFourn.ColWidths [GridFourn.colcount-2] := 0;
        GridFourn.ColWidths [GridFourn.colcount-3] := 0;
        GridFourn.ColWidths [GridFourn.colcount-4] := 0;
        END;
     end;
  {$ENDIF}
  until critere='';
  //
  BZoom := TToolbarButton97(ecran.FindComponent('BZOOM'));
  BZoom.onclick := BZoom_OnClick;
  //
End ;


// Pour afficher un look up des fournisseurs d'un article
procedure TOF_GridFournisseur.OnLoad ;
Var Messages: TControl;
GridMessage: THGrid;
RechFourn: String ;
i,j,nbcaractere: Integer ;
Trouver: boolean ;
begin
  inherited ;
  nbcaractere:=Length(rang)+1 ;
  Trouver:=False ;
  if Typaff = '2'
    then Ecran.Caption:=' Références Fournisseurs non affectées à un article '
    else Ecran.Caption:=' Fournisseurs de l''article ';
  Messages:=GetControl('GRIDFOURN') ;
  GridMessage:=THGrid(Messages) ;
  if LaTob <> NIL then
  begin
       If rang<>'' then
       begin
         While not Trouver do
            begin
            nbcaractere:=nbcaractere-1 ;
            j:=0;
            for i:=0 to LaTob.Detail.count-1 do
            begin
              RechFourn:=copy(LaTob.Detail[j].GetValue('GCA_TIERS'),1,nbcaractere);
              rang:=copy(rang,1,nbcaractere) ;
              if (compareStr(RechFourn,rang) = 0) then
              begin
                   Trouver:=True ;
                   break ;
              end ;
              j:=j+1 ;
            end;
         end;
         LaTob.PutGridDetail(GridMessage,False,False,LesColonnes,False) ;
         GridMessage.Row:=j+1 ;
       end else
       LaTob.PutGridDetail(GridMessage,False,False,LesColonnes,False) ;
  end;
TFVierge (ecran).hmtrad.resizegridcolumns(gridmessage);
//LaTob:=Nil;
end;

function TOFGridFournisseur_OnInsert(Parms : Array of variant; nb: integer) : variant ;
var  F : TForm ;
     TheTOF : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFvierge) then TheTof:=TFVierge(F).LaTOF else exit;
  if (TheTof is TOF_GridFournisseur) then
  begin
    TOF_GridFournisseur(TheTof).Insert;
  end else exit;
end;

{$IFDEF BTP}
function BTPRefreshLaLigne (Parms : Array of variant; nb: integer) : variant;
var  F : TForm ;
     TheTOF : TOF ;
     LaLigne : integer;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFvierge) then TheTof:=TFVierge(F).LaTOF else exit;
  if (TheTof is TOF_GridFournisseur) then
  begin
    TOF_GridFournisseur(TheTof).BTPRefreshLaLigne;
  end else exit;
end;  
{$ENDIF}

procedure InitTOFGridFournisseur ;
begin
RegisterAglFunc('OnInsert', True , 1, TOFGridFournisseur_OnInsert) ;
{$IFDEF BTP}
RegisterAglFunc('BTPRefreshLaLigne', True , 1, BTPRefreshLaLigne) ;
{$ENDIF}
End ;

Initialization
registerclasses([TOF_GridFournisseur]) ;
InitTOFGridFournisseur ;
end.

