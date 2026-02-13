import { NextRequest, NextResponse } from "next/server";
import { getAuthUrl } from "@/lib/graph-client";
import { isAuthenticated } from "@/lib/token-store";

export async function GET(request: NextRequest) {
  const host = request.headers.get("host") || undefined;

  // Check if already authenticated
  if (isAuthenticated()) {
    return NextResponse.json({
      success: true,
      authenticated: true,
      message: "Already authenticated with Microsoft Graph",
    });
  }

  // Return the auth URL for the client to redirect to
  const authUrl = getAuthUrl(host);
  return NextResponse.json({
    success: true,
    authenticated: false,
    authUrl,
  });
}
