import JSZip from 'jszip';
import {
  DEFAULT_ANDROID_CONFIG,
  generateActivityMainXml,
  generateAdaptiveForegroundXml,
  generateAdaptiveIconXml,
  generateAndroidManifestXml,
  generateAppBuildGradleKts,
  generateColorsXml,
  generateFilePathsXml,
  generateGitHubActionsWorkflow,
  generateGradleProperties,
  generateGradleWrapperProperties,
  generateMainActivityKt,
  generateProjectReadme,
  generateRootBuildGradleKts,
  generateSettingsGradleKts,
  generateStringsXml,
  generateThemesXml,
  type AndroidAppConfig,
} from './androidConfig';
import { getApiBaseOrigin } from './googleAuth';

export interface SourceFileItem {
  path: string;
  category: string;
  size: number;
  content: string;
}

export function buildAndroidSourceFileItems(
  incoming?: Partial<AndroidAppConfig>
): SourceFileItem[] {
  const config: AndroidAppConfig = {
    ...DEFAULT_ANDROID_CONFIG,
    ...incoming,
    versionCode: Number(incoming?.versionCode ?? DEFAULT_ANDROID_CONFIG.versionCode),
  };
  const pkgPath = config.packageName.replace(/\./g, '/');
  const rawList: Array<{ path: string; category: string; content: string }> = [
    { path: 'README.md', category: 'Root & Gradle', content: generateProjectReadme(config) },
    {
      path: 'settings.gradle.kts',
      category: 'Root & Gradle',
      content: generateSettingsGradleKts(config),
    },
    { path: 'build.gradle.kts', category: 'Root & Gradle', content: generateRootBuildGradleKts() },
    { path: 'gradle.properties', category: 'Root & Gradle', content: generateGradleProperties() },
    {
      path: 'gradle/wrapper/gradle-wrapper.properties',
      category: 'Root & Gradle',
      content: generateGradleWrapperProperties(),
    },
    {
      path: '.github/workflows/build-apk.yml',
      category: 'GitHub Actions',
      content: generateGitHubActionsWorkflow(config),
    },
    { path: 'app/build.gradle.kts', category: 'Android App', content: generateAppBuildGradleKts(config) },
    { path: 'app/proguard-rules.pro', category: 'Android App', content: '# Custom ProGuard rules\n' },
    {
      path: 'app/src/main/AndroidManifest.xml',
      category: 'Android App',
      content: generateAndroidManifestXml(config),
    },
    {
      path: `app/src/main/java/${pkgPath}/MainActivity.kt`,
      category: 'Android Kotlin',
      content: generateMainActivityKt(config),
    },
    {
      path: 'app/src/main/res/layout/activity_main.xml',
      category: 'Android Resources',
      content: generateActivityMainXml(),
    },
    {
      path: 'app/src/main/res/values/strings.xml',
      category: 'Android Resources',
      content: generateStringsXml(config),
    },
    {
      path: 'app/src/main/res/values/colors.xml',
      category: 'Android Resources',
      content: generateColorsXml(config),
    },
    {
      path: 'app/src/main/res/values/themes.xml',
      category: 'Android Resources',
      content: generateThemesXml(),
    },
    {
      path: 'app/src/main/res/xml/file_paths.xml',
      category: 'Android Resources',
      content: generateFilePathsXml(),
    },
    {
      path: 'app/src/main/res/mipmap-anydpi-v26/ic_launcher.xml',
      category: 'Android Resources',
      content: generateAdaptiveIconXml(),
    },
    {
      path: 'app/src/main/res/mipmap-anydpi-v26/ic_launcher_round.xml',
      category: 'Android Resources',
      content: generateAdaptiveIconXml(),
    },
    {
      path: 'app/src/main/res/drawable/ic_launcher_foreground.xml',
      category: 'Android Resources',
      content: generateAdaptiveForegroundXml(config),
    },
  ];
  return rawList.map((f) => ({
    ...f,
    size: new TextEncoder().encode(f.content).length,
  }));
}

export async function fetchWebSourceFilesWithFallback(): Promise<SourceFileItem[]> {
  const apiBaseOrigin = getApiBaseOrigin();
  const tryFetchJson = async (url: string): Promise<SourceFileItem[] | null> => {
    try {
      const res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) return null;
      const ct = (res.headers.get('content-type') || '').toLowerCase();
      if (ct.includes('text/html')) return null;
      const data = (await res.json()) as {
        ok?: boolean;
        files?: SourceFileItem[];
      };
      if (Array.isArray(data.files) && data.files.length > 0) {
        return data.files;
      }
      return null;
    } catch {
      return null;
    }
  };

  const liveFiles = await tryFetchJson(`${apiBaseOrigin}/api/source-files`);
  if (liveFiles) return liveFiles;

  const snapshotFiles = await tryFetchJson('/source-snapshot.json');
  if (snapshotFiles) return snapshotFiles;

  throw new Error('সোর্স ফাইল লোড করা যায়নি। ইন্টারনেট বা সার্ভার সংযোগ পরীক্ষা করুন।');
}

export async function fetchAndroidSourceFilesWithFallback(
  appName?: string,
  packageName?: string
): Promise<SourceFileItem[]> {
  const apiBaseOrigin = getApiBaseOrigin();
  const configPatch: Partial<AndroidAppConfig> = {
    ...DEFAULT_ANDROID_CONFIG,
    appName: appName?.trim() || DEFAULT_ANDROID_CONFIG.appName,
    packageName: packageName?.trim() || DEFAULT_ANDROID_CONFIG.packageName,
    targetUrl: typeof window !== 'undefined' ? window.location.origin : DEFAULT_ANDROID_CONFIG.targetUrl,
  };
  try {
    const res = await fetch(`${apiBaseOrigin}/api/android-source-files`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(configPatch),
    });
    const ct = (res.headers.get('content-type') || '').toLowerCase();
    if (res.ok && !ct.includes('text/html')) {
      const data = (await res.json()) as { ok?: boolean; files?: SourceFileItem[] };
      if (data.ok && Array.isArray(data.files) && data.files.length > 0) {
        return data.files;
      }
    }
  } catch {
    // Fall back to client-side generation
  }
  return buildAndroidSourceFileItems(configPatch);
}

export async function buildSourceZipClientFallback(
  scope: 'full' | 'web',
  appName?: string,
  packageName?: string
): Promise<Blob> {
  const webFiles = await fetchWebSourceFilesWithFallback();
  const zip = new JSZip();
  for (const f of webFiles) {
    zip.file(f.path, f.content);
  }
  if (scope === 'full') {
    const androidFiles = await fetchAndroidSourceFilesWithFallback(appName, packageName);
    for (const f of androidFiles) {
      zip.file(`android/${f.path}`, f.content);
    }
    const wf = androidFiles.find((f) => f.path === '.github/workflows/build-apk.yml');
    if (wf) {
      zip.file('.github/workflows/build-apk.yml', wf.content);
    }
  }
  return zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

export async function buildFilesZipFromList(files: SourceFileItem[]): Promise<Blob> {
  const zip = new JSZip();
  for (const f of files) {
    zip.file(f.path, f.content);
  }
  return zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}
