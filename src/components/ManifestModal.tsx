import React, { useState } from 'react';
import {
  Download,
  Copy,
  Check,
  X,
  FileCode,
  HelpCircle,
  ExternalLink,
  ShieldCheck,
  FolderSync,
} from 'lucide-react';

interface ManifestModalProps {
  isOpen: boolean;
  onClose: () => void;
  appUrl: string;
}

export const ManifestModal: React.FC<ManifestModalProps> = ({
  isOpen,
  onClose,
  appUrl,
}) => {
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState<'manifest' | 'instructions'>('instructions');

  if (!isOpen) return null;

  // The base URL where this add-in is hosted
  const normalizedUrl = appUrl.endsWith('/') ? appUrl.slice(0, -1) : appUrl;

  const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<!-- Manifiesto oficial para Microsoft Word 2016 Web Add-in -->
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">

  <!-- Identificador único del complemento -->
  <Id>c1b4a4d1-f382-47bf-877e-90af682986d0</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Word Add-in Systems</ProviderName>
  <DefaultLocale>es-ES</DefaultLocale>
  <DisplayName DefaultValue="Segmentador de Bloques y Semáforos Word 2016"/>
  <Description DefaultValue="Panel lateral de Word 2016 para segmentación en bloques virtuales y calificación con semáforo excluyente de Datos, Experiencia y Tareas."/>
  
  <IconUrl DefaultValue="${normalizedUrl}/favicon.ico"/>
  <HighResolutionIconUrl DefaultValue="${normalizedUrl}/favicon.ico"/>
  <SupportUrl DefaultValue="${normalizedUrl}"/>

  <AppDomains>
    <AppDomain>${normalizedUrl}</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Document"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="${normalizedUrl}"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <!-- Compatibilidad con cinta de opciones de Word 2016 Desktop -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="TabHome">
              <Group id="WordBlockGroup">
                <Label resid="GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon16"/>
                  <bt:Image size="32" resid="Icon32"/>
                  <bt:Image size="80" resid="Icon80"/>
                </Icon>
                <Control xsi:type="Button" id="OpenTaskpaneButton">
                  <Label resid="ButtonLabel"/>
                  <Supertip>
                    <Title resid="ButtonTitle"/>
                    <Description resid="ButtonDesc"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon16"/>
                    <bt:Image size="32" resid="Icon32"/>
                    <bt:Image size="80" resid="Icon80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="TaskpaneUrl"/>
                  </Action>
                </Control>
              </Group>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon16" DefaultValue="${normalizedUrl}/favicon.ico"/>
        <bt:Image id="Icon32" DefaultValue="${normalizedUrl}/favicon.ico"/>
        <bt:Image id="Icon80" DefaultValue="${normalizedUrl}/favicon.ico"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="TaskpaneUrl" DefaultValue="${normalizedUrl}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Semáforo Bloques"/>
        <bt:String id="ButtonLabel" DefaultValue="Panel de Bloques"/>
        <bt:String id="ButtonTitle" DefaultValue="Segmentador de Bloques"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="ButtonDesc" DefaultValue="Abre el panel lateral de evaluación y semáforo de secciones del documento."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;

  const handleDownload = () => {
    const blob = new Blob([manifestXml], { type: 'application/xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'manifest-word2016-bloques.xml';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(manifestXml);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full max-h-[90vh] flex flex-col overflow-hidden border border-slate-200">
        {/* Header */}
        <div className="bg-[#2b579a] text-white p-4 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <FileCode className="w-5 h-5" />
            <div>
              <h2 className="text-sm font-bold tracking-tight">
                Instalación en Microsoft Word 2016 (Manifiesto XML)
              </h2>
              <p className="text-[11px] text-blue-100">
                Archivo de manifiesto listo para cargar directamente en tu Word 2016 de escritorio
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg hover:bg-white/20 text-white transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab switcher */}
        <div className="flex border-b border-slate-200 bg-slate-50 px-4 text-xs font-semibold">
          <button
            onClick={() => setActiveTab('instructions')}
            className={`py-2.5 px-3 border-b-2 flex items-center gap-1.5 transition-colors ${
              activeTab === 'instructions'
                ? 'border-[#2b579a] text-[#2b579a]'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <HelpCircle className="w-3.5 h-3.5" />
            <span>Guía Paso a Paso para Word 2016</span>
          </button>
          <button
            onClick={() => setActiveTab('manifest')}
            className={`py-2.5 px-3 border-b-2 flex items-center gap-1.5 transition-colors ${
              activeTab === 'manifest'
                ? 'border-[#2b579a] text-[#2b579a]'
                : 'border-transparent text-slate-500 hover:text-slate-800'
            }`}
          >
            <FileCode className="w-3.5 h-3.5" />
            <span>Ver Código XML del Manifiesto</span>
          </button>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-5 text-xs text-slate-700">
          {activeTab === 'instructions' ? (
            <div className="space-y-4">
              <div className="bg-blue-50 border border-blue-200 p-3.5 rounded-xl text-blue-900 space-y-1">
                <div className="font-bold flex items-center gap-1.5">
                  <ShieldCheck className="w-4 h-4 text-[#2b579a]" />
                  <span>Método Oficial de Microsoft para Word 2016 (Sideloading sin tienda)</span>
                </div>
                <p className="text-[11px] text-blue-800">
                  En Word 2016 para Windows, los complementos web se instalan en 1 minuto configurando un catálogo de carpetas compartidas local. No necesitas cuenta empresarial de Microsoft ni subir nada a la tienda.
                </p>
              </div>

              <ol className="space-y-3.5 list-none">
                <li className="flex gap-3">
                  <span className="w-6 h-6 rounded-full bg-[#2b579a] text-white flex items-center justify-center font-bold text-xs shrink-0">
                    1
                  </span>
                  <div>
                    <strong className="text-slate-900 block mb-0.5">Descarga el archivo manifest.xml</strong>
                    <p className="text-slate-600 text-[11px]">
                      Haz clic en el botón inferior "Descargar manifest.xml" y guárdalo en una carpeta en tu disco local (por ejemplo en: <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">C:\OfficeAddins\</code>).
                    </p>
                  </div>
                </li>

                <li className="flex gap-3">
                  <span className="w-6 h-6 rounded-full bg-[#2b579a] text-white flex items-center justify-center font-bold text-xs shrink-0">
                    2
                  </span>
                  <div>
                    <strong className="text-slate-900 block mb-0.5">Comparte la carpeta en Windows</strong>
                    <p className="text-slate-600 text-[11px]">
                      En el Explorador de Windows, haz clic derecho sobre la carpeta <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">OfficeAddins</code> &rarr; <strong>Propiedades</strong> &rarr; pestaña <strong>Compartir</strong> &rarr; botón <strong>Compartir...</strong> &rarr; añade tu usuario con permisos de lectura &rarr; Clic en Compartir. Copia la ruta de red generada (ej: <code className="bg-slate-100 px-1 py-0.5 rounded font-mono">\\MI-PC\OfficeAddins</code>).
                    </p>
                  </div>
                </li>

                <li className="flex gap-3">
                  <span className="w-6 h-6 rounded-full bg-[#2b579a] text-white flex items-center justify-center font-bold text-xs shrink-0">
                    3
                  </span>
                  <div>
                    <strong className="text-slate-900 block mb-0.5">Configura Word 2016</strong>
                    <p className="text-slate-600 text-[11px]">
                      Abre Word 2016 &rarr; Menú <strong>Archivo</strong> &rarr; <strong>Opciones</strong> &rarr; <strong>Centro de confianza</strong> &rarr; <strong>Configuración del Centro de confianza...</strong> &rarr; <strong>Catálogos de complementos de confianza</strong>.
                      Pega la ruta de red (<code className="bg-slate-100 px-1 py-0.5 rounded font-mono">\\MI-PC\OfficeAddins</code>) &rarr; Pulsa <strong>Agregar catálogo</strong> &rarr; Marca la casilla <strong>"Mostrar en el menú"</strong> &rarr; Pulsa Aceptar.
                    </p>
                  </div>
                </li>

                <li className="flex gap-3">
                  <span className="w-6 h-6 rounded-full bg-[#2b579a] text-white flex items-center justify-center font-bold text-xs shrink-0">
                    4
                  </span>
                  <div>
                    <strong className="text-slate-900 block mb-0.5">Abre el complemento en Word</strong>
                    <p className="text-slate-600 text-[11px]">
                      Reinicia Word 2016 &rarr; Ve a la pestaña <strong>Insertar</strong> &rarr; <strong>Mis complementos</strong> &rarr; Haz clic en la pestaña <strong>CARPETA COMPARTIDA</strong> &rarr; Selecciona <strong>"Segmentador de Bloques y Semáforos"</strong>. ¡El panel lateral se abrirá inmediatamente a la derecha!
                    </p>
                  </div>
                </li>
              </ol>
            </div>
          ) : (
            <div className="space-y-2">
              <div className="flex justify-between items-center">
                <span className="text-[11px] text-slate-500 font-mono">
                  URL configurada: {normalizedUrl}
                </span>
                <button
                  onClick={handleCopy}
                  className="px-2 py-1 bg-slate-100 hover:bg-slate-200 rounded font-semibold text-slate-700 text-[11px] flex items-center gap-1"
                >
                  {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copied ? 'Copiado al portapapeles' : 'Copiar XML'}</span>
                </button>
              </div>
              <pre className="p-3 bg-slate-900 text-slate-100 rounded-xl font-mono text-[11px] overflow-x-auto max-h-72">
                {manifestXml}
              </pre>
            </div>
          )}
        </div>

        {/* Footer Actions */}
        <div className="bg-slate-100 border-t border-slate-200 p-3.5 flex items-center justify-between">
          <div className="text-[11px] text-slate-500">
            Compatible con Word 2016, 2019, 2021 y Microsoft 365
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopy}
              className="px-3 py-1.5 border border-slate-300 bg-white hover:bg-slate-50 rounded-lg text-slate-700 font-semibold text-xs flex items-center gap-1.5 transition-colors"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copied ? 'Copiado' : 'Copiar XML'}</span>
            </button>
            <button
              onClick={handleDownload}
              className="px-4 py-1.5 bg-[#2b579a] hover:bg-[#204377] text-white rounded-lg font-semibold text-xs flex items-center gap-1.5 shadow-sm transition-all"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Descargar manifest.xml</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
