"use client";
import React, { useState, useEffect, useMemo, useRef } from "react";
import { Tree, Upload, Button, Input, Tag, Empty, Spin, Alert, Modal } from "antd";
import {
  PictureOutlined,
  CloudServerOutlined,
  FolderOutlined,
  AppstoreOutlined,
  SearchOutlined,
  CheckCircleFilled,
  LoginOutlined,
  LogoutOutlined,
  SafetyCertificateOutlined,
  ExclamationCircleOutlined,
  FileTextOutlined,
  BankOutlined,
} from "@ant-design/icons";
import type { TreeDataNode } from "antd";
import { useSession } from "next-auth/react";
import { useRouter } from "next/navigation";
import { msalInstance, ARM_SCOPES, ensureMsalInitialized } from "@/lib/msal";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

// Helper to update a specific node in the tree with new children
function updateTreeData(list: TreeDataNode[], key: React.Key, children: TreeDataNode[]): TreeDataNode[] {
  return list.map((node) => {
    if (node.key === key) {
      return { ...node, children };
    }
    if (node.children) {
      return { ...node, children: updateTreeData(node.children, key, children) };
    }
    return node;
  });
}

// ─── Helper: count all leaf nodes under a tree ───────────────────────────────
function countLeaves(nodes: TreeDataNode[]): number {
  let count = 0;
  for (const node of nodes) {
    if (node.isLeaf) count++;
    else if (node.children) count += countLeaves(node.children);
  }
  return count;
}

// ─── Helper: filter tree nodes by search query ───────────────────────────────
function filterTree(nodes: TreeDataNode[], query: string): TreeDataNode[] {
  if (!query) return nodes;
  const lower = query.toLowerCase();
  return nodes
    .map((node) => {
      const titleStr = String(node.title).toLowerCase();
      const childMatches = node.children ? filterTree(node.children, query) : [];
      if (titleStr.includes(lower) || childMatches.length > 0) {
        return { ...node, children: childMatches.length > 0 ? childMatches : node.children };
      }
      return null;
    })
    .filter(Boolean) as TreeDataNode[];
}

// ─── Helper: collect all keys for expanding ──────────────────────────────────
function collectAllKeys(nodes: TreeDataNode[]): React.Key[] {
  const keys: React.Key[] = [];
  for (const node of nodes) {
    keys.push(node.key);
    if (node.children) keys.push(...collectAllKeys(node.children));
  }
  return keys;
}

export default function FinalConfigPage() {
  const { data: session, status } = useSession();

  // ─── Azure (client-tenant) connection state via MSAL ───────────────────────
  const [armToken, setArmToken] = useState<string | null>(null);
  const [connectedTenant, setConnectedTenant] = useState<{
    tenantId: string;
    username: string;
  } | null>(null);
  const [connecting, setConnecting] = useState(false);
  const [msalReady, setMsalReady] = useState(false);
  // Guest access: optional client tenant (domain or tenant ID) + the authority used for the current connection
  const [tenantInput, setTenantInput] = useState("");
  const [authority, setAuthority] = useState<string | undefined>(undefined);

  const [treeData, setTreeData] = useState<TreeDataNode[]>([]);
  const [loadingInitial, setLoadingInitial] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [checkedKeys, setCheckedKeys] = useState<React.Key[]>([]);
  const [expandedKeys, setExpandedKeys] = useState<React.Key[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedLogo, setSelectedLogo] = useState<File | null>(null);
  const [logoPreviewUrl, setLogoPreviewUrl] = useState<string | null>(null);
  const [projectTitle, setProjectTitle] = useState("");
  const [clientName, setClientName] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const subsFetchedRef = useRef(false);
  // Guard: prevents double-fire (React StrictMode / rapid clicks) that causes
  // the "interaction_in_progress" MSAL error.
  const connectingRef = useRef(false);

  const router = useRouter();

  const handleGenerate = async () => {
    if (!armToken || checkedKeys.length === 0 || !projectTitle || !clientName) return;

    setIsGenerating(true);
    setError(null);
    try {
      // 6.1: get a fresh ARM token (MSAL returns the cached one or refreshes it silently)
      let freshToken = armToken;
      try {
        const account = msalInstance.getActiveAccount();
        if (account) {
          const result = await msalInstance.acquireTokenSilent({
            scopes: ARM_SCOPES,
            account,
            ...(authority ? { authority } : {}),
          });
          freshToken = result.accessToken;
          setArmToken(freshToken);
        }
      } catch {
        throw new Error("Your Azure session has expired. Please disconnect and connect to the client tenant again.");
      }

      const formData = new FormData();
      formData.append("projectTitle", projectTitle);
      formData.append("clientName", clientName);
      formData.append("tenantId", connectedTenant?.tenantId || "");
      formData.append("armToken", freshToken || "");

      const leafKeys = checkedKeys.filter((k) => String(k).toLowerCase().startsWith("/subscriptions"));
      formData.append("checkedKeys", JSON.stringify(leafKeys));

      if (selectedLogo) {
        formData.append("logo", selectedLogo);
      }

      const res = await fetch(`${API_BASE_URL}/api/final-config/generate`, {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const errData = await res.json().catch(() => null);
        const detail = typeof errData?.detail === "string" ? errData.detail : null;
        throw new Error(detail || `Failed to start generation (HTTP ${res.status}).`);
      }

      const data = await res.json();
      if (data.task_id) {
        router.push(`/generate/final-config/tasks/${data.task_id}/processing`);
      }
    } catch (err: any) {
      setError(err.message || "An error occurred starting the generation.");
      setIsGenerating(false);
    }
  };

  const totalResources = useMemo(() => countLeaves(treeData), [treeData]);
  const allKeys = useMemo(() => collectAllKeys(treeData), [treeData]);
  const filteredTreeData = useMemo(() => filterTree(treeData, searchQuery), [treeData, searchQuery]);

  // ─── Initialize MSAL ONCE on mount + clear any stuck interaction state ──────
  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        await ensureMsalInitialized();
        // CRITICAL: resolves/clears any leftover interaction flag from a
        // previous interrupted popup (the real cause of interaction_in_progress).
        await msalInstance.handleRedirectPromise();

        // Restore an existing session if the user already connected earlier.
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          msalInstance.setActiveAccount(accounts[0]);
        }
      } catch (e) {
        // handleRedirectPromise can throw if nothing to handle — safe to ignore.
        console.debug("[MSAL init]", e);
      } finally {
        if (mounted) setMsalReady(true);
      }
    })();
    return () => {
      mounted = false;
    };
  }, []);

  // ─── Connect to Client Tenant (MSAL popup, /organizations authority) ───────
  const connectToClientTenant = async () => {
    // Hard guard against concurrent interactions
    if (connectingRef.current) return;
    connectingRef.current = true;

    setError(null);
    setConnecting(true);
    try {
      await ensureMsalInitialized();

      // Guest access (Scenario 2): sign in to the client's tenant instead of the home tenant.
      // Empty field = same behaviour as before (/organizations -> home tenant).
      const tenant = tenantInput.trim();
      const loginAuthority = tenant ? `https://login.microsoftonline.com/${tenant}` : undefined;
      const authorityOpt = loginAuthority ? { authority: loginAuthority } : {};

      // Helper that runs the popup; retries once if MSAL reports a stuck interaction.
      const doLoginPopup = async () => {
        try {
          return await msalInstance.loginPopup({
            scopes: ARM_SCOPES,
            prompt: "select_account",
            ...authorityOpt,
          });
        } catch (e: any) {
          const m = String(e?.errorCode || e?.message || e);
          if (m.includes("interaction_in_progress")) {
            // Clear the stuck state and try one more time.
            await msalInstance.clearCache();
            await new Promise((r) => setTimeout(r, 300));
            return await msalInstance.loginPopup({
              scopes: ARM_SCOPES,
              prompt: "select_account",
              ...authorityOpt,
            });
          }
          throw e;
        }
      };

      const loginResult = await doLoginPopup();

      if (!loginResult.account) {
        throw new Error("No account returned from login.");
      }
      msalInstance.setActiveAccount(loginResult.account);

      // Acquire the ARM token silently for the chosen account/tenant
      const tokenResult = await msalInstance.acquireTokenSilent({
        scopes: ARM_SCOPES,
        account: loginResult.account,
        ...authorityOpt,
      });

      setArmToken(tokenResult.accessToken);
      setAuthority(loginAuthority);
      setConnectedTenant({
        // The token's tenant = the tenant we are really connected to (the client's tenant for guests)
        tenantId: tokenResult.tenantId || loginResult.account.tenantId,
        username: loginResult.account.username,
      });

      // Reset the tree so we freshly load the newly-connected tenant
      subsFetchedRef.current = false;
      setTreeData([]);
      setCheckedKeys([]);
      setExpandedKeys([]);
    } catch (err: any) {
      const msg = String(err?.errorCode || err?.message || err);
      if (
        msg.includes("AADSTS65001") || // consent required
        msg.includes("AADSTS90094") || // admin consent required
        msg.toLowerCase().includes("admin_consent") ||
        msg.toLowerCase().includes("consent_required")
      ) {
        setError(
          "This client tenant requires an administrator to approve access. " +
          "Please ask the client's Azure administrator to grant admin consent for this application, then try again."
        );
      } else if (
        msg.toLowerCase().includes("user_cancelled") ||
        msg.toLowerCase().includes("popup_window_error")
      ) {
        setError("Sign-in was cancelled or the popup was blocked. Please allow popups and try again.");
      } else if (msg.includes("interaction_in_progress")) {
        setError("A previous sign-in is still finishing. Please refresh the page and try again.");
      } else {
        setError(`Failed to connect to client tenant: ${msg}`);
      }
    } finally {
      setConnecting(false);
      connectingRef.current = false;
    }
  };

  // ─── Disconnect (clear the client-tenant session) ──────────────────────────
  const disconnectClientTenant = async () => {
    try {
      await msalInstance.clearCache();
    } catch {
      /* non-critical */
    }
    setArmToken(null);
    setConnectedTenant(null);
    setAuthority(undefined);
    setTreeData([]);
    setCheckedKeys([]);
    setExpandedKeys([]);
    subsFetchedRef.current = false;
    setError(null);
  };

  // Ask before disconnecting, since it clears the tree and all selections
  const confirmDisconnect = () => {
    Modal.confirm({
      title: "Disconnect from this tenant?",
      icon: <ExclamationCircleOutlined />,
      content: "Your selected resources will be cleared. You will need to connect again to continue.",
      okText: "Disconnect",
      okType: "danger",
      cancelText: "Cancel",
      onOk: disconnectClientTenant,
    });
  };

  // ─── Fetch subscriptions once we have an ARM token ─────────────────────────
  useEffect(() => {
    if (!armToken) return;
    if (subsFetchedRef.current) return;
    subsFetchedRef.current = true;

    async function fetchSubscriptions() {
      setLoadingInitial(true);
      try {
        const res = await fetch(`${API_BASE_URL}/api/azure/subscriptions`, {
          headers: { Authorization: `Bearer ${armToken}` },
        });
        if (!res.ok) throw new Error(`Failed to fetch subscriptions: ${res.statusText}`);
        const data = await res.json();

        const nodes: TreeDataNode[] = data.map((sub: any) => ({
          title: `${sub.displayName} (${sub.subscriptionId})`,
          key: sub.subscriptionId,
          icon: <CloudServerOutlined style={{ color: "#1677ff" }} />,
          isLeaf: false,
          children: undefined, // forces lazy load on expand
        }));

        setTreeData(nodes);
      } catch (err: any) {
        setError(err.message || "Failed to load subscriptions");
      } finally {
        setLoadingInitial(false);
      }
    }

    fetchSubscriptions();
  }, [armToken]);

  // ─── Lazy Load: Resource Groups & Resources on expand ──────────────────────
  const onLoadData = async (node: any): Promise<React.Key[]> => {
    if (node.children) return []; // already loaded

    const keyParts = String(node.key).split("|");
    const isSubscription = keyParts.length === 1;

    try {
      if (isSubscription) {
        const subId = keyParts[0];
        const res = await fetch(`${API_BASE_URL}/api/azure/subscriptions/${subId}/resourcegroups`, {
          headers: { Authorization: `Bearer ${armToken}` },
        });
        if (!res.ok) throw new Error("Failed to load resource groups");
        const data = await res.json();

        const rgNodes: TreeDataNode[] = data.map((rg: any) => ({
          title: rg.name,
          key: `${subId}|${rg.name}`,
          icon: <FolderOutlined style={{ color: "#722ed1" }} />,
          isLeaf: false,
        }));

        setTreeData((origin) => updateTreeData(origin, node.key, rgNodes.length ? rgNodes : []));
        return rgNodes.map(n => n.key);
      } else if (keyParts.length === 2) {
        const subId = keyParts[0];
        const rgName = keyParts[1];

        const res = await fetch(
          `${API_BASE_URL}/api/azure/subscriptions/${subId}/resourcegroups/${rgName}/resources`,
          { headers: { Authorization: `Bearer ${armToken}` } }
        );
        if (!res.ok) throw new Error("Failed to load resources");
        const data = await res.json();

        const resourceNodes: TreeDataNode[] = data.map((r: any) => ({
          title: `${r.name} (${r.type.split("/").pop()})`,
          key: r.id, // Full Azure resource ID
          icon: <AppstoreOutlined style={{ color: "#52c41a" }} />,
          isLeaf: true,
        }));

        setTreeData((origin) => updateTreeData(origin, node.key, resourceNodes.length ? resourceNodes : []));
        return resourceNodes.map(n => n.key);
      }
    } catch (err: any) {
      console.error(err);
      setTreeData((origin) => updateTreeData(origin, node.key, []));
    }
    return [];
  };

  // Auto-expand all when searching
  useEffect(() => {
    if (searchQuery) setExpandedKeys(allKeys);
  }, [searchQuery, allKeys]);

  // Logo preview URL lifecycle
  useEffect(() => {
    if (selectedLogo) {
      const objectUrl = URL.createObjectURL(selectedLogo);
      setLogoPreviewUrl(objectUrl);
      return () => URL.revokeObjectURL(objectUrl);
    } else {
      setLogoPreviewUrl(null);
    }
  }, [selectedLogo]);

  const checkedLeafCount = checkedKeys.filter((k) => String(k).toLowerCase().startsWith("/subscriptions")).length;
  // Project details are mandatory before generating
  const projectDetailsFilled = projectTitle.trim() !== "" && clientName.trim() !== "";
  const canGenerate = checkedLeafCount > 0 && projectDetailsFilled;
  // Only show the form sections when the connected tenant has at least one subscription
  const hasSubscriptions = treeData.length > 0;

  // ─── Gate: must be logged into the app (Enfrasys) first ────────────────────
  const appLoggedIn = status === "authenticated" && !!session;

  return (
    <div className="container-fluid bg-light d-flex flex-column flex-grow-1 align-items-center py-4 pb-5">
      {/* ── Title Section ── */}
      <div className="text-center mt-4 mb-4">
        <h2 className="fw-bold mb-2">Generate Final Configuration Document</h2>
        <p className="text-muted mb-0">
          Connect to a client tenant, then select Azure resources to generate the final configuration document.
        </p>
      </div>

      {error && (
        <Alert
          title="Notice"
          description={error}
          type="error"
          showIcon
          closable
          onClose={() => setError(null)}
          className="mb-4 w-100"
          style={{ maxWidth: "720px" }}
        />
      )}

      {/* ── Client Tenant Connection Section ── */}
      <div className="w-100 mb-4" style={{ maxWidth: "720px" }}>
        <div className="bg-white border rounded p-4 shadow-sm">
          <div className="d-flex align-items-center justify-content-between">
            <div className="d-flex align-items-center">
              <div
                className="rounded d-flex align-items-center justify-content-center me-3"
                style={{
                  width: "48px",
                  height: "48px",
                  fontSize: "22px",
                  color: "#fff",
                  background: connectedTenant
                    ? "linear-gradient(135deg, #52c41a 0%, #73d13d 100%)"
                    : "linear-gradient(135deg, #1677ff 0%, #4096ff 100%)",
                }}
              >
                <SafetyCertificateOutlined />
              </div>
              <div>
                <h6 className="fw-semibold mb-1">Client Tenant Connection</h6>
                {connectedTenant ? (
                  <p className="text-muted small mb-0">
                    Connected as <span className="fw-medium">{connectedTenant.username}</span>
                    <br />
                    Tenant ID: <code style={{ fontSize: "11px" }}>{connectedTenant.tenantId}</code>
                  </p>
                ) : (
                  <p className="text-muted small mb-0">
                    Sign in to the client&apos;s Azure tenant to browse their resources.
                  </p>
                )}
              </div>
            </div>

            {connectedTenant ? (
              <Button icon={<LogoutOutlined />} onClick={confirmDisconnect} danger>
                Disconnect
              </Button>
            ) : (
              <Button
                type="primary"
                icon={<LoginOutlined />}
                loading={connecting}
                disabled={!appLoggedIn || !msalReady}
                onClick={connectToClientTenant}
              >
                Connect to Client Tenant
              </Button>
            )}
          </div>

          {!connectedTenant && (
            <div className="mt-3">
              <label className="form-label text-muted small fw-semibold mb-1">
                Client Tenant <span className="fw-normal">(optional, required for guest access)</span>
              </label>
              <Input
                placeholder="e.g. contoso.com or 00000000-0000-0000-0000-000000000000"
                value={tenantInput}
                onChange={(e) => setTenantInput(e.target.value)}
                allowClear
              />
              <p className="text-muted mt-1 mb-0" style={{ fontSize: "12px" }}>
                Leave empty to use your account&apos;s own tenant. Enter the client&apos;s domain or tenant ID if you were invited as a guest.
              </p>
            </div>
          )}

          {!appLoggedIn && (
            <Alert
              className="mt-3"
              type="warning"
              showIcon
              title="You must be signed in to the application first before connecting to a client tenant."
            />
          )}
        </div>
      </div>

      {/* ── Azure Resource Tree Section (only after connected) ── */}
      {connectedTenant && (
        <div className="w-100" style={{ maxWidth: "720px" }}>
          <div className="bg-white border rounded shadow-sm" style={{ overflow: "hidden" }}>
            {/* Tree Header Bar */}
            <div
              className="d-flex align-items-center justify-content-between px-4 py-3"
              style={{
                borderBottom: "1px solid #f0f0f0",
                background: "linear-gradient(135deg, #fafbff 0%, #f5f7ff 100%)",
              }}
            >
              <div className="d-flex align-items-center gap-2">
                <div
                  className="d-flex align-items-center justify-content-center rounded"
                  style={{
                    width: "36px",
                    height: "36px",
                    background: "linear-gradient(135deg, #1677ff 0%, #4096ff 100%)",
                    fontSize: "18px",
                    color: "#fff",
                  }}
                >
                  <CloudServerOutlined />
                </div>
                <div>
                  <span className="fw-semibold" style={{ fontSize: "14px" }}>
                    Azure Resources
                  </span>
                  <div className="d-flex align-items-center gap-2 mt-1">
                    <Tag color="blue" style={{ margin: 0, fontSize: "11px" }}>
                      {treeData.length} Subscriptions
                    </Tag>
                    {checkedLeafCount > 0 && (
                      <Tag color="green" icon={<CheckCircleFilled />} style={{ margin: 0, fontSize: "11px" }}>
                        {checkedLeafCount} {totalResources > 0 ? `/ ${totalResources}` : ""} selected
                      </Tag>
                    )}
                  </div>
                </div>
              </div>

              <div className="d-flex gap-2">
                <Button
                  size="small"
                  type="text"
                  style={{ fontSize: "12px", color: "#1677ff" }}
                  onClick={() => setExpandedKeys(allKeys)}
                >
                  Expand All
                </Button>
                <Button
                  size="small"
                  type="text"
                  style={{ fontSize: "12px", color: "#8c8c8c" }}
                  onClick={() => setExpandedKeys([])}
                >
                  Collapse All
                </Button>
              </div>
            </div>

            {/* Search Bar */}
            <div className="px-4 pt-3 pb-2">
              <Input
                placeholder="Search subscriptions, resource groups, or resources..."
                prefix={<SearchOutlined style={{ color: "#bfbfbf" }} />}
                allowClear
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ borderRadius: "8px" }}
              />
            </div>

            {/* Tree Component */}
            <div className="px-4 pt-2 pb-4" style={{ maxHeight: "420px", overflowY: "auto" }}>
              {loadingInitial ? (
                <div className="d-flex justify-content-center py-5">
                  <Spin tip="Loading Subscriptions..." size="large" />
                </div>
              ) : filteredTreeData.length > 0 ? (
                <Tree
                  checkable
                  showIcon
                  selectable={false}
                  loadData={onLoadData}
                  treeData={filteredTreeData}
                  checkedKeys={checkedKeys}
                  expandedKeys={expandedKeys}
                  onExpand={(keys) => setExpandedKeys(keys)}
                  onCheck={async (keys, e: any) => {
                    let currentChecked = Array.isArray(keys) ? keys : keys.checked;
                    setCheckedKeys(currentChecked);

                    if (e.checked && !e.node.isLeaf && (!e.node.children || e.node.children.length === 0)) {
                      setExpandedKeys((prev) => Array.from(new Set([...prev, e.node.key])));

                      const loadedKeys = await onLoadData(e.node);
                      if (loadedKeys && loadedKeys.length > 0) {
                        setCheckedKeys((prev) => Array.from(new Set([...prev, ...loadedKeys])));

                        // If it's a Subscription, we just loaded RGs. Now recursively load resources for these RGs.
                        const isSubscription = String(e.node.key).indexOf("|") === -1;
                        if (isSubscription) {
                          let allResourceKeys: React.Key[] = [];
                          setExpandedKeys((prev) => Array.from(new Set([...prev, ...loadedKeys])));

                          const promises = loadedKeys.map((rgKey) =>
                            onLoadData({ key: rgKey, children: undefined })
                          );

                          const results = await Promise.all(promises);
                          results.forEach((resKeys) => {
                            if (resKeys) allResourceKeys.push(...resKeys);
                          });

                          if (allResourceKeys.length > 0) {
                            setCheckedKeys((prev) => Array.from(new Set([...prev, ...allResourceKeys])));
                          }
                        }
                      }
                    }
                  }}
                  style={{ fontSize: "13.5px" }}
                />
              ) : (
                <Empty
                  image={Empty.PRESENTED_IMAGE_SIMPLE}
                  description={
                    <span className="text-muted" style={{ fontSize: "13px" }}>
                      {searchQuery ? (
                        `No resources match "${searchQuery}"`
                      ) : (
                        <>
                          No subscriptions are available to this account in this tenant.
                          <br />
                          Ask the client to assign the <strong>Reader</strong> role on the subscription or resource groups to be documented, then disconnect and connect again.
                        </>
                      )}
                    </span>
                  }
                  style={{ padding: "40px 0" }}
                />
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Project Details Section (only after connected, with subscriptions) ── */}
      {connectedTenant && hasSubscriptions && (
        <div className="bg-white border rounded p-4 mt-4 shadow-sm w-100" style={{ maxWidth: "720px" }}>
          <div className="d-flex align-items-center mb-3">
            <div
              className="bg-light border rounded d-flex align-items-center justify-content-center me-3"
              style={{ width: "48px", height: "48px", fontSize: "22px" }}
            >
              <FileTextOutlined className="text-secondary" />
            </div>
            <div>
              <h6 className="fw-semibold mb-1">Project Details</h6>
              {/* <p className="text-muted small mb-0">Optional &bull; Used in the generated document title page and headers</p> */}
              <p className="text-muted small mb-0">Required &bull; Used in the generated document title page and headers</p>
            </div>
          </div>

          <div className="row g-3 mt-1">
            <div className="col-12 col-md-6">
              <label className="form-label text-muted small fw-semibold mb-1">
                Project Title <span className="text-danger">*</span>
              </label>
              <Input
                size="large"
                placeholder="e.g. Cloud Migration Phase 2"
                value={projectTitle}
                onChange={(e) => setProjectTitle(e.target.value)}
                status={checkedLeafCount > 0 && !projectTitle.trim() ? "error" : undefined}
                prefix={<FileTextOutlined className="text-muted" style={{ marginRight: 4 }} />}
              />
            </div>
            <div className="col-12 col-md-6">
              <label className="form-label text-muted small fw-semibold mb-1">
                Client Company Name <span className="text-danger">*</span>
              </label>
              <Input
                size="large"
                placeholder="e.g. Contoso Ltd"
                value={clientName}
                onChange={(e) => setClientName(e.target.value)}
                status={checkedLeafCount > 0 && !clientName.trim() ? "error" : undefined}
                prefix={<BankOutlined className="text-muted" style={{ marginRight: 4 }} />}
              />
            </div>
          </div>

          <p className="text-muted mt-3 mb-0" style={{ fontSize: "12px" }}>
            These values will appear on the document cover page and headers.
          </p>
        </div>
      )}

      {/* ── Upload Client Logo Section (only after connected, with subscriptions) ── */}
      {connectedTenant && hasSubscriptions && (
        <div className="bg-white border rounded p-4 mt-4 shadow-sm w-100" style={{ maxWidth: "720px" }}>
          <div className="d-flex align-items-center justify-content-between">
            <div className="d-flex align-items-center">
              <div
                className="bg-light border rounded d-flex align-items-center justify-content-center me-3"
                style={{ width: "48px", height: "48px", fontSize: "22px" }}
              >
                <PictureOutlined className="text-secondary" />
              </div>
              <div>
                <h6 className="fw-semibold mb-1">Upload Client Logo</h6>
                <p className="text-muted small mb-0">Optional &bull; png, jpeg, jpg</p>
              </div>
            </div>

            <Upload
              name="logo"
              accept="image/png, image/jpeg, image/jpg"
              showUploadList={false}
              beforeUpload={(file: any) => {
                setSelectedLogo(file);
                return false;
              }}
            >
              <Button size="middle">{selectedLogo ? "Select another logo" : "Select"}</Button>
            </Upload>
          </div>

          {logoPreviewUrl && selectedLogo && (
            <div
              className="mt-4 pt-4 border-top d-flex flex-column align-items-center"
              style={{ animation: "fadeIn 0.3s ease-in-out" }}
            >
              <p
                className="text-muted small mb-3 fw-semibold text-uppercase"
                style={{ letterSpacing: "0.5px" }}
              >
                Logo Preview
              </p>
              <img
                src={logoPreviewUrl}
                alt="Client Logo Preview"
                style={{ maxHeight: "140px", maxWidth: "100%", objectFit: "contain" }}
                className="border rounded p-3 bg-light mb-3 shadow-sm"
              />
              <span className="text-secondary small fw-medium">{selectedLogo.name}</span>
            </div>
          )}
        </div>
      )}

      {/* ── Generate Button (only after connected, with subscriptions) ── */}
      {connectedTenant && hasSubscriptions && (
        <div className="w-100 mt-4 d-flex justify-content-center" style={{ maxWidth: "720px" }}>
          {/* <Button
            type="primary"
            size="large"
            disabled={checkedLeafCount === 0}
            style={{
              minWidth: "240px",
              height: "48px",
              borderRadius: "8px",
              fontWeight: 600,
              fontSize: "15px",
            }}
          >
            {checkedLeafCount > 0
              ? `Generate Configuration (${checkedLeafCount} resource${checkedLeafCount > 1 ? "s" : ""})`
              : "Select resources to generate"}
          </Button> */}
          <Button
            type="primary"
            size="large"
            loading={isGenerating}
            disabled={!canGenerate || isGenerating}
            onClick={handleGenerate}
            style={{
              minWidth: "240px",
              height: "48px",
              borderRadius: "8px",
              fontWeight: 600,
              fontSize: "15px",
            }}
          >
            {checkedLeafCount === 0
              ? "Select resources to generate"
              : !projectDetailsFilled
                ? "Fill in project details to generate"
                : `Generate Configuration (${checkedLeafCount} resource${checkedLeafCount > 1 ? "s" : ""})`}
          </Button>
        </div>
      )}
    </div>
  );
}

