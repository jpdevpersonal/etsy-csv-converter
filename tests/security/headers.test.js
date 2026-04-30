const { getHeadersForPath, parseCspDirectives } = require("../helpers/headers-config");

describe("privacy and security header configuration", () => {
  test("root path inherits the documented security headers", () => {
    const headers = getHeadersForPath("/");
    const csp = parseCspDirectives(headers["Content-Security-Policy"]);

    expect(headers["Referrer-Policy"]).toBe("no-referrer");
    expect(headers["X-Content-Type-Options"]).toBe("nosniff");
    expect(headers["X-Frame-Options"]).toBe("DENY");
    expect(headers["Permissions-Policy"]).toContain("camera=()");
    expect(headers["Permissions-Policy"]).toContain("microphone=()");
    expect(headers["Strict-Transport-Security"]).toContain("max-age=31536000");

    expect(csp["default-src"]).toEqual(["'self'"]);
    expect(csp["base-uri"]).toEqual(["'self'"]);
    expect(csp["object-src"]).toEqual(["'none'"]);
    expect(csp["frame-ancestors"]).toEqual(["'none'"]);
    expect(csp["form-action"]).toEqual(["'self'"]);
    expect(csp["script-src"]).toContain("'self'");
    expect(csp["style-src"]).toContain("'self'");
    expect(csp["img-src"]).toContain("data:");
    expect(csp["connect-src"]).toContain("'self'");
    expect(csp["manifest-src"]).toEqual(["'self'"]);
    expect(csp["media-src"]).toEqual(["'self'"]);
    expect(csp["worker-src"]).toContain("'self'");
    expect(csp["upgrade-insecure-requests"]).toEqual([]);
  });

  test("configured CSP keeps the page local-first while documenting current inline exceptions", () => {
    const headers = getHeadersForPath("/");
    const csp = parseCspDirectives(headers["Content-Security-Policy"]);

    expect(csp["script-src"]).toContain("'unsafe-inline'");
    expect(csp["style-src"]).toContain("'unsafe-inline'");
    expect(csp["script-src"]).toContain("https://pagead2.googlesyndication.com");
    expect(csp["frame-src"]).toContain("https://tpc.googlesyndication.com");
    expect(csp["connect-src"]).toContain("https://www.googleadservices.com");
  });
});

const deployedBaseUrl = process.env.DEPLOYED_BASE_URL;
const maybeTestDeployedHeaders = deployedBaseUrl ? test : test.skip;

maybeTestDeployedHeaders("deployed site serves the expected privacy and security headers", async () => {
  const response = await fetch(deployedBaseUrl, { redirect: "manual" });
  const csp = parseCspDirectives(response.headers.get("content-security-policy"));

  expect(response.headers.get("referrer-policy")).toBe("no-referrer");
  expect(response.headers.get("x-content-type-options")).toBe("nosniff");
  expect(response.headers.get("x-frame-options")).toBe("DENY");
  expect(response.headers.get("permissions-policy") || "").toContain("camera=()");
  expect(response.headers.get("strict-transport-security") || "").toContain("max-age=31536000");
  expect(csp["default-src"]).toContain("'self'");
  expect(csp["frame-ancestors"]).toContain("'none'");
  expect(csp["object-src"]).toContain("'none'");
});