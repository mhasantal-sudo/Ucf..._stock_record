export interface GitHubFileEntry {
  path: string;
  base64Content: string;
}

export interface GitHubPushOptions {
  token: string;
  repository: string;
  branch?: string;
  commitMessage?: string;
  autoCreateRepo?: boolean;
  isPrivateRepo?: boolean;
  cleanTree?: boolean;
  files: GitHubFileEntry[];
  onProgress?: (phase: string, done: number, total: number) => void;
}

export interface GitHubPushResult {
  ok: true;
  owner: string;
  repo: string;
  branch: string;
  repoUrl: string;
  commitUrl: string;
  actionsUrl: string;
  commitSha: string;
  filesUploaded: number;
  createdRepo: boolean;
}

const GITHUB_API_BASE = 'https://api.github.com';

export function utf8ToBase64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  return uint8ArrayToBase64(bytes);
}

export function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

export function parseGitHubRepositoryInput(rawInput: string): { owner: string; repo: string } {
  const cleaned = rawInput
    .trim()
    .replace(/^https?:\/\/(www\.)?github\.com\//i, '')
    .replace(/^git@github\.com:/i, '')
    .replace(/\.git$/i, '')
    .replace(/^\/+|\/+$/g, '');

  if (!cleaned) {
    throw new Error('GitHub রিপোজিটরির নাম বা লিংক দিন (যেমন: khaddo-boraddo অথবা username/khaddo-boraddo)।');
  }

  const parts = cleaned.split('/').filter(Boolean);
  if (parts.length === 1) {
    return { owner: '', repo: parts[0] };
  }
  if (parts.length >= 2) {
    return { owner: parts[0], repo: parts[1] };
  }
  throw new Error('সঠিক GitHub রিপোজিটরির নাম দিন (যেমন: username/repo-name)।');
}

async function ghRequest<T>(
  token: string,
  path: string,
  init: RequestInit = {}
): Promise<{ status: number; data: T; headers: Headers }> {
  const res = await fetch(`${GITHUB_API_BASE}${path}`, {
    ...init,
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${token}`,
      'X-GitHub-Api-Version': '2022-11-28',
      ...(init.body ? { 'Content-Type': 'application/json' } : {}),
      ...(init.headers || {}),
    },
  });

  const text = await res.text();
  let data: unknown = {};
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { message: text };
    }
  }

  return { status: res.status, data: data as T, headers: res.headers };
}

function formatGitHubErrorMessage(status: number, data: { message?: string } | undefined, fallback: string): string {
  const rawMsg = (data?.message || '').trim();
  const lower = rawMsg.toLowerCase();

  if (status === 401 || lower.includes('bad credentials')) {
    return 'GitHub টোকেনটি সঠিক নয় অথবা মেয়াদ শেষ হয়েছে। নতুন টোকেন তৈরি করে পেস্ট করুন।';
  }
  if (lower.includes('workflow') && (lower.includes('scope') || lower.includes('permission') || status === 422 || status === 404)) {
    return 'টোকেনে "workflow" অনুমতি (scope) দেওয়া নেই! উপরের «১ ক্লিকে Token তৈরি করুন» বাটন দিয়ে repo এবং workflow দুটোই টিক দিয়ে নতুন টোকেন নিন।';
  }
  if (status === 403 && lower.includes('rate limit')) {
    return 'GitHub API রেট লিমিট পূর্ণ হয়েছে। কয়েক মিনিট পর আবার চেষ্টা করুন।';
  }
  if (status === 403) {
    return `GitHub অনুমতি প্রত্যাখ্যান করেছে: ${rawMsg || 'টোকেনে repo ও workflow scope আছে কিনা দেখুন।'}`;
  }
  return rawMsg ? `${fallback} (${rawMsg})` : fallback;
}

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export async function pushFilesToGitHub(options: GitHubPushOptions): Promise<GitHubPushResult> {
  const token = options.token.trim();
  if (!token) {
    throw new Error('GitHub Personal Access Token দিন।');
  }
  if (!Array.isArray(options.files) || options.files.length === 0) {
    throw new Error('আপলোড করার মতো কোনো সোর্স ফাইল পাওয়া যায়নি।');
  }

  const parsedRepo = parseGitHubRepositoryInput(options.repository);
  const autoCreateRepo = options.autoCreateRepo !== false;
  const isPrivateRepo = Boolean(options.isPrivateRepo);
  const commitMessage =
    options.commitMessage?.trim() || 'Update Khaddo Boraddo web & Android source code';

  // Deduplicate files cleanly by normalized path
  const fileMap = new Map<string, GitHubFileEntry>();
  for (const item of options.files) {
    const cleanPath = item.path.replace(/^\/+/, '').replace(/\/+/g, '/');
    if (!cleanPath || cleanPath.endsWith('/')) continue;
    fileMap.set(cleanPath, { path: cleanPath, base64Content: item.base64Content });
  }
  const files = Array.from(fileMap.values());
  const hasWorkflowFiles = files.some((f) => f.path.startsWith('.github/workflows/'));

  options.onProgress?.('GitHub অ্যাকাউন্ট যাচাই করা হচ্ছে...', 0, files.length);

  // 1. Verify authenticated user & scopes
  const userRes = await ghRequest<{ login?: string; message?: string }>(token, '/user');
  if (userRes.status >= 400 || !userRes.data?.login) {
    throw new Error(
      formatGitHubErrorMessage(userRes.status, userRes.data, 'GitHub অ্যাকাউন্ট যাচাই করা যায়নি')
    );
  }
  const authLogin = userRes.data.login;

  // Classic PATs expose X-OAuth-Scopes; Fine-grained PATs omit it
  const oauthScopesHeader = userRes.headers.get('x-oauth-scopes');
  if (oauthScopesHeader) {
    const scopes = oauthScopesHeader
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter(Boolean);
    if (!scopes.includes('repo') && !scopes.includes('public_repo')) {
      throw new Error(
        'টোকেনে "repo" অনুমতি নেই। উপরের «১ ক্লিকে Token তৈরি করুন» বাটনে ট্যাপ করে repo ও workflow টিক দিয়ে নতুন টোকেন তৈরি করুন।'
      );
    }
    if (hasWorkflowFiles && !scopes.includes('workflow')) {
      throw new Error(
        'টোকেনে "workflow" অনুমতি নেই (.github/workflows/build-apk.yml পুশ করতে workflow টিক দেওয়া আবশ্যক)। উপরের «১ ক্লিকে Token তৈরি করুন» লিংক ব্যবহার করুন।'
      );
    }
  }

  const owner = parsedRepo.owner || authLogin;
  const repo = parsedRepo.repo;

  // 2. Check if repository exists; create if requested
  options.onProgress?.(`রিপোজিটরি (${owner}/${repo}) পরীক্ষা করা হচ্ছে...`, 0, files.length);
  let createdRepo = false;
  let repoRes = await ghRequest<{
    name?: string;
    default_branch?: string;
    html_url?: string;
    message?: string;
  }>(token, `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}`);

  if (repoRes.status === 404) {
    if (!autoCreateRepo) {
      throw new Error(
        `GitHub-এ "${owner}/${repo}" রিপোজিটরি পাওয়া যায়নি। অটো-তৈরি অপশন চালু করুন অথবা আগে GitHub-এ রিপোজিটরি বানান।`
      );
    }
    if (owner.toLowerCase() !== authLogin.toLowerCase()) {
      throw new Error(
        `"${owner}/${repo}" রিপোজিটরি পাওয়া যায়নি এবং অন্য ইউজার/অর্গানাইজেশনের নামে অটো-তৈরি করা যাবে না।`
      );
    }

    options.onProgress?.(`নতুন রিপোজিটরি (${owner}/${repo}) তৈরি করা হচ্ছে...`, 0, files.length);
    const createRes = await ghRequest<{
      name?: string;
      default_branch?: string;
      html_url?: string;
      message?: string;
    }>(token, '/user/repos', {
      method: 'POST',
      body: JSON.stringify({
        name: repo,
        private: isPrivateRepo,
        auto_init: true,
        description: 'খাদ্যশস্য বরাদ্দ ও বিতরণ ব্যবস্থাপনা ওয়েব ও অ্যান্ড্রয়েড অ্যাপ',
      }),
    });
    if (createRes.status >= 400 || !createRes.data?.name) {
      throw new Error(
        formatGitHubErrorMessage(createRes.status, createRes.data, 'নতুন রিপোজিটরি তৈরি করা যায়নি')
      );
    }
    createdRepo = true;
    repoRes = createRes;
    await delay(1200);
  } else if (repoRes.status >= 400) {
    throw new Error(
      formatGitHubErrorMessage(repoRes.status, repoRes.data, 'রিপোজিটরি তথ্য পাওয়া যায়নি')
    );
  }

  const defaultBranch = repoRes.data?.default_branch || 'main';
  const targetBranch = options.branch?.trim() || defaultBranch || 'main';
  const repoPath = `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}`;

  // 3. Resolve target branch head commit (and initialize if repo is completely empty)
  let refRes = await ghRequest<{ object?: { sha: string }; message?: string }>(
    token,
    `${repoPath}/git/ref/heads/${encodeURIComponent(targetBranch)}`
  );

  if (refRes.status === 409 || refRes.status === 404) {
    const branchesRes = await ghRequest<Array<{ name: string }>>(token, `${repoPath}/branches`);
    const branches = Array.isArray(branchesRes.data) ? branchesRes.data : [];

    if (branches.length === 0) {
      // Completely empty repo: seed initial commit via Contents API so Git Data API works
      options.onProgress?.('খালি রিপোজিটরি প্রস্তুত করা হচ্ছে...', 0, files.length);
      const seedRes = await ghRequest<{ message?: string }>(token, `${repoPath}/contents/README.md`, {
        method: 'PUT',
        body: JSON.stringify({
          message: 'Initialize repository',
          content: utf8ToBase64('# Khaddo Boraddo App\n'),
          branch: targetBranch,
        }),
      });
      if (seedRes.status >= 400) {
        throw new Error(
          formatGitHubErrorMessage(seedRes.status, seedRes.data, 'খালি রিপোজিটরি ইনিশিয়ালাইজ করা যায়নি')
        );
      }
      await delay(800);
      refRes = await ghRequest<{ object?: { sha: string }; message?: string }>(
        token,
        `${repoPath}/git/ref/heads/${encodeURIComponent(targetBranch)}`
      );
    } else if (refRes.status === 404) {
      // Target branch doesn't exist yet; fork from default branch
      const baseRefRes = await ghRequest<{ object?: { sha: string }; message?: string }>(
        token,
        `${repoPath}/git/ref/heads/${encodeURIComponent(branches[0].name)}`
      );
      if (baseRefRes.status >= 400 || !baseRefRes.data?.object?.sha) {
        throw new Error(
          formatGitHubErrorMessage(baseRefRes.status, baseRefRes.data, 'মূল ব্রাঞ্চ পাওয়া যায়নি')
        );
      }
      const createRefRes = await ghRequest<{ object?: { sha: string }; message?: string }>(
        token,
        `${repoPath}/git/refs`,
        {
          method: 'POST',
          body: JSON.stringify({
            ref: `refs/heads/${targetBranch}`,
            sha: baseRefRes.data.object.sha,
          }),
        }
      );
      if (createRefRes.status >= 400 || !createRefRes.data?.object?.sha) {
        throw new Error(
          formatGitHubErrorMessage(createRefRes.status, createRefRes.data, 'নতুন ব্রাঞ্চ তৈরি করা যায়নি')
        );
      }
      refRes = createRefRes;
    }
  }

  if (refRes.status >= 400 || !refRes.data?.object?.sha) {
    throw new Error(
      formatGitHubErrorMessage(refRes.status, refRes.data, `ব্রাঞ্চ (${targetBranch}) অ্যাক্সেস করা যায়নি`)
    );
  }

  const headCommitSha = refRes.data.object.sha;
  const headCommitRes = await ghRequest<{ tree?: { sha: string }; message?: string }>(
    token,
    `${repoPath}/git/commits/${headCommitSha}`
  );
  if (headCommitRes.status >= 400 || !headCommitRes.data?.tree?.sha) {
    throw new Error(
      formatGitHubErrorMessage(headCommitRes.status, headCommitRes.data, 'সর্বশেষ কমিট তথ্য পড়া যায়নি')
    );
  }
  const baseTreeSha = headCommitRes.data.tree.sha;

  // 4. Upload all file blobs in concurrent batches
  const uploadedBlobs: Array<{ path: string; sha: string }> = new Array(files.length);
  const CONCURRENCY = 8;
  let nextIdx = 0;
  let completedCount = 0;

  const uploadWorker = async () => {
    while (nextIdx < files.length) {
      const idx = nextIdx++;
      const file = files[idx];

      let lastErr: Error | null = null;
      for (let attempt = 0; attempt < 2; attempt++) {
        const blobRes = await ghRequest<{ sha?: string; message?: string }>(
          token,
          `${repoPath}/git/blobs`,
          {
            method: 'POST',
            body: JSON.stringify({
              content: file.base64Content,
              encoding: 'base64',
            }),
          }
        );
        if (blobRes.status < 400 && blobRes.data?.sha) {
          uploadedBlobs[idx] = { path: file.path, sha: blobRes.data.sha };
          lastErr = null;
          break;
        }
        lastErr = new Error(
          formatGitHubErrorMessage(blobRes.status, blobRes.data, `ফাইল আপলোড ব্যর্থ: ${file.path}`)
        );
        if (blobRes.status >= 500 || blobRes.status === 429) {
          await delay(500);
          continue;
        }
        break;
      }

      if (lastErr) throw lastErr;
      completedCount++;
      options.onProgress?.('ফাইল আপলোড হচ্ছে...', completedCount, files.length);
    }
  };

  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, files.length) }, () => uploadWorker()));

  // 5. Create Git Tree atomically
  options.onProgress?.('Git ট্রি ও কমিট তৈরি হচ্ছে...', files.length, files.length);
  const treePayload: {
    tree: Array<{ path: string; mode: '100644'; type: 'blob'; sha: string }>;
    base_tree?: string;
  } = {
    tree: uploadedBlobs.map((b) => ({
      path: b.path,
      mode: '100644',
      type: 'blob',
      sha: b.sha,
    })),
  };
  if (!options.cleanTree) {
    treePayload.base_tree = baseTreeSha;
  }

  const treeRes = await ghRequest<{ sha?: string; message?: string }>(token, `${repoPath}/git/trees`, {
    method: 'POST',
    body: JSON.stringify(treePayload),
  });
  if (treeRes.status >= 400 || !treeRes.data?.sha) {
    throw new Error(
      formatGitHubErrorMessage(treeRes.status, treeRes.data, 'Git ট্রি তৈরি করা যায়নি')
    );
  }

  // 6. Create Git Commit
  const commitRes = await ghRequest<{ sha?: string; html_url?: string; message?: string }>(
    token,
    `${repoPath}/git/commits`,
    {
      method: 'POST',
      body: JSON.stringify({
        message: commitMessage,
        tree: treeRes.data.sha,
        parents: [headCommitSha],
      }),
    }
  );
  if (commitRes.status >= 400 || !commitRes.data?.sha) {
    throw new Error(
      formatGitHubErrorMessage(commitRes.status, commitRes.data, 'Git কমিট তৈরি করা যায়নি')
    );
  }

  // 7. Update Branch Reference
  const updateRefRes = await ghRequest<{ message?: string }>(
    token,
    `${repoPath}/git/refs/heads/${encodeURIComponent(targetBranch)}`,
    {
      method: 'PATCH',
      body: JSON.stringify({
        sha: commitRes.data.sha,
        force: false,
      }),
    }
  );
  if (updateRefRes.status >= 400) {
    throw new Error(
      formatGitHubErrorMessage(updateRefRes.status, updateRefRes.data, 'ব্রাঞ্চ আপডেট করা যায়নি')
    );
  }

  const repoUrl = `https://github.com/${owner}/${repo}`;
  return {
    ok: true,
    owner,
    repo,
    branch: targetBranch,
    repoUrl,
    commitUrl: commitRes.data.html_url || `${repoUrl}/commit/${commitRes.data.sha}`,
    actionsUrl: `${repoUrl}/actions`,
    commitSha: commitRes.data.sha,
    filesUploaded: files.length,
    createdRepo,
  };
}


