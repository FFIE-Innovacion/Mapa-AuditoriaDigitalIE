# Solicitud a TI · Acceso de lectura a la Matriz FFIE vía Entra ID

**Solicitante:** Innovación UG FFIE (innovacion@ffie.com.co)
**Propósito:** que un proceso automático (GitHub Actions, una vez al día a las 6:00 am) descargue el archivo
**Matriz FFIE.xlsx** del sitio de SharePoint **SRVINNOVACIN2** para alimentar los dashboards de Auditoría Digital IE.
El proceso solo **lee** el archivo; no escribe, no borra, no toca nada más del tenant.

## Qué necesitamos que TI cree

### 1. App Registration en Microsoft Entra ID

| Campo | Valor |
|---|---|
| Nombre | `FFIE-Innovacion-Dashboard-Sync` |
| Tipos de cuenta | Solo esta organización (single tenant) |
| Redirect URI | Ninguna (es un proceso sin interfaz, flujo *client credentials*) |

### 2. Client secret

- Certificates & secrets → New client secret.
- Vigencia sugerida: **24 meses** (o la máxima que permita la política). Por favor indicarnos la fecha de vencimiento
  para agendar la renovación.

### 3. Permiso de API — mínimo privilegio

Microsoft Graph → **Application permissions** → **`Sites.Selected`** → *Grant admin consent*.

`Sites.Selected` **no da acceso a ningún sitio por sí solo**. El acceso se otorga después, sitio por sitio, con el paso 4.
Esto es preferible a `Sites.Read.All`, que permitiría leer todo SharePoint del tenant.

### 4. Otorgar a la app permiso de *lectura* sobre el sitio SRVINNOVACIN2

Con PnP PowerShell (como administrador de SharePoint):

```powershell
Connect-PnPOnline -Url https://ffie3.sharepoint.com/sites/SRVINNOVACIN2 -Interactive
Grant-PnPAzureADAppSitePermission -AppId "<CLIENT_ID de la app>" -DisplayName "FFIE-Innovacion-Dashboard-Sync" -Permissions Read -Site https://ffie3.sharepoint.com/sites/SRVINNOVACIN2
```

O vía Microsoft Graph (con una identidad que tenga `Sites.FullControl.All`):

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
{
  "roles": ["read"],
  "grantedToIdentities": [{ "application": { "id": "<CLIENT_ID>", "displayName": "FFIE-Innovacion-Dashboard-Sync" } }]
}
```

## Qué necesitamos que TI nos entregue

| Dato | Dónde se ve en Entra ID |
|---|---|
| **Tenant ID** | App Registration → Overview → *Directory (tenant) ID* |
| **Client ID** | App Registration → Overview → *Application (client) ID* |
| **Client Secret** | El valor generado en el paso 2 (solo se muestra una vez) |
| Fecha de vencimiento del secret | Para agendar la renovación |

Los tres valores se guardan como **GitHub Secrets** (cifrados) en el repositorio del dashboard; nunca quedan en código.
El proceso usa el flujo OAuth 2.0 *client credentials* contra `login.microsoftonline.com` y llama únicamente a:

- `GET /sites/ffie3.sharepoint.com:/sites/SRVINNOVACIN2` (resolver el sitio)
- `GET /sites/{id}/drive/root:/<ruta>` y `.../content` (leer el archivo)

## Renovación del secret

Cuando el secret venza, el job fallará con un error claro de autenticación y avisará por correo de GitHub.
Bastará con generar un secret nuevo en la misma app y reemplazar el valor en GitHub Secrets. No hay que rehacer nada más.
