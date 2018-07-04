unit UMenusNeg;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     Hent1,
		 Menus,
		 HCtrls,
     HMsgBox,
     SysUtils,
     AglInit,
     forms,
     UtilPGI,
     EntGC,
     uTob;


function GestionMenuNegoce (Numtag : integer) : integer;
function GestionMenuParcMateriel(Numtag : integer) : integer;

implementation

function GestionMenuNegoce (Numtag : integer) : integer;
begin

	if not VH_GC.SeriaNEGOCE then
  begin
  	PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Commerce de détail');
    result := -1;
    exit;
  end;

  result := 0;

  Case Numtag of
    // --- DEVIS
    328110 : AGLLanceFiche('BTP','BTDEVISNEG_MUL','GP_NATUREPIECEG=DE;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    328120 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=DE','','VEN;VIVANTE') ;
    328130 : AGLLanceFiche('BTP','BTACCDEVNEG_MUL','GP_NATUREPIECEG=DE','','MODIFICATION') ;    // Acceptation de Devis
    328140 : AGLLanceFiche('BTP','BTREJDEVNEG_MUL','GP_NATUREPIECEG=DE;AFF_ETATAFFAIRE=ENC','','MODIFICATION') ;   // Rejet de Devis
    // --- COMMANDES
    328210 : AGLLanceFiche('BTP','BTPIECENEG_MUL','GP_NATUREPIECEG=CC;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    328220 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=CC','','VEN;VIVANTE') ;
    // --- LIVRAISONS
    328310 : AGLLanceFiche('BTP','BTTRANSPIECE_MUL','GP_NATUREPIECEG=CC;GP_VIVANTE=X','','BLC') ;  // Livraisons
   	328311 : AGLLanceFiche('BTP','BTGROUPEMANNEG','GP_VENTEACHAT=VEN','','LIVRAISONCLI;VENTE;;BLC') ;  // Comm --> Livr
    328320 : AGLLanceFiche('BTP','BTPIECENEG_MUL','GP_NATUREPIECEG=BLC;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    328330 : AGLLanceFiche('BTP','BTRESTEALIVRER','','','');
    328340 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=BLC','','VEN;VIVANTE') ;
    // --- Factures
    328410 : AGLLanceFiche('BTP','BTSELFACNEG_MUL','GL_NATUREPIECEG=BLC','','NEXTPIECE=') ;
    328420 : AGLLanceFiche('BTP','BTRESTEAFAC','','','') ;


    else HShowMessage('2;?caption?;'+'Fonction non disponible : '+';W;O;O;O;','Commerce de détail',IntToStr(Numtag)) ;
  end;
end;

function GestionMenuParcMateriel(Numtag : integer) : integer;
begin

	if not VH_GC.BTSeriaParcMateriel then
  begin
  	PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Gestion Parc Matériel');
    result := -1;
    exit;
  end;

  result := 0;

  Case Numtag of
    // --- DEVIS
    329110 : AGLLanceFiche('BTP','BTFAMILLEMAT_MUL','','','');
    329120 : AGLLanceFiche('BTP','BTMATERIEL_MUL','','','');
    329130 : AGLLanceFiche('BTP','BTTYPEACTIONMAT_MUL','','','');
  else HShowMessage('2;?caption?;'+'Fonction non disponible : '+';W;O;O;O;','Gestion Parc Matériels',IntToStr(Numtag)) ;
  end;


end;


end.
