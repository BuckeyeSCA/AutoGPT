"use client";

import { useQueryClient } from "@tanstack/react-query";

import {
  getGetV2ListChatConnectionsQueryKey,
  useGetV2ListChatConnections,
  useGetV2ListProviderModelTiers,
} from "@/app/api/__generated__/endpoints/chat/chat";
import { useOAuthConnect } from "@/components/contextual/IntegrationsPanel/components/ConnectServiceDialog/components/DetailView/useOAuthConnect";

import { useOnboardingWizardStore } from "../../store";
import {
  hasLinkedSubscription,
  linkedModelsSentence,
  linkedSubscriptionName,
} from "./helpers";

export function useConnectStep() {
  const queryClient = useQueryClient();
  const nextStep = useOnboardingWizardStore((s) => s.nextStep);

  const connectionsQuery = useGetV2ListChatConnections({
    query: { refetchOnWindowFocus: false },
  });
  const offers =
    connectionsQuery.data?.status === 200
      ? connectionsQuery.data.data.offers
      : undefined;

  // What ChatGPT's tiers resolve to, which the connections list cannot answer
  // for a connection the user has not made yet.
  const tiersQuery = useGetV2ListProviderModelTiers({
    query: { refetchOnWindowFocus: false },
  });
  const providers =
    tiersQuery.data?.status === 200
      ? tiersQuery.data.data.providers
      : undefined;

  function finishConnection() {
    queryClient.invalidateQueries({
      queryKey: getGetV2ListChatConnectionsQueryKey(),
    });
    nextStep();
  }

  const { connect, isPending } = useOAuthConnect({
    provider: "codex",
    onSuccess: finishConnection,
  });

  return {
    connect,
    finishConnection,
    isConnecting: isPending,
    skip: nextStep,
    isAlreadyLinked: hasLinkedSubscription(offers),
    linkedSubscription: linkedSubscriptionName(offers),
    models: linkedModelsSentence(providers),
  };
}
