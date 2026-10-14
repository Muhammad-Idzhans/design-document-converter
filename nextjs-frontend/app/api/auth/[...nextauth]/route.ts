import NextAuth from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";

const clientId = process.env.ENFRASYS_APPLICATION_CLIENT_ID;
const clientSecret = process.env.ENFRASYS_CLIENT_SECRET_VALUE;
const tenantId = process.env.ENFRASYS_DIRECTORY_TENANT_ID;

const providers = [];
if (clientId && clientSecret) {
  providers.push(
    AzureADProvider({
      clientId,
      clientSecret,
      tenantId,
      authorization: {
        params: {
          scope: "openid profile email https://management.azure.com/user_impersonation",
        },
      },
    })
  );
}

const handler = NextAuth({
  providers,
  session: {
    strategy: "jwt",
  },
  pages: {
    signIn: "/", // Redirect to our custom login page if unauthenticated
  },
  callbacks: {
    async jwt({ token, profile, account }) {
      if (profile) {
        token.name = profile.name;
        // Azure AD sends email as "preferred_username", not "email"
        token.email = (profile as any).preferred_username || (profile as any).email || token.email;
      }
      if (account) {
        token.accessToken = account.access_token;
      }
      return token;
    },
    async session({ session, token }) {
      if (session.user) {
        session.user.name = token.name as string;
        session.user.email = token.email as string;
      }
      (session as any).accessToken = token.accessToken;
      return session;
    },
  },
});

export { handler as GET, handler as POST };
