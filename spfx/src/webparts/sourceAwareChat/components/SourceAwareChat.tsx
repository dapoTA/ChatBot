import * as React from 'react';
import styles from './SourceAwareChat.module.scss';

export interface ISourceAwareChatProps {
  chatbotUrl: string;
  loginName: string;
  displayName: string;
}

interface IKnowledgeSource {
  id: number;
  name: string;
  description: string;
  isPortalWide: boolean;
}

interface IChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
}

function normaliseServerUrl(value: string): string {
  return String(value || 'https://chatbot.technicalassurance.com').replace(/\/+$/, '');
}

function makeSessionId(): string {
  return `spfx-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

const SourceAwareChat: React.FC<ISourceAwareChatProps> = ({
  chatbotUrl,
  loginName,
  displayName
}) => {
  const serverUrl = React.useMemo(() => normaliseServerUrl(chatbotUrl), [chatbotUrl]);
  const [sources, setSources] = React.useState<IKnowledgeSource[]>([]);
  const [selectedSourceId, setSelectedSourceId] = React.useState<string>('');
  const [question, setQuestion] = React.useState<string>('');
  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [isLoadingSources, setIsLoadingSources] = React.useState<boolean>(true);
  const [isSending, setIsSending] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const sessionId = React.useRef<string>(makeSessionId());

  React.useEffect(() => {
    let cancelled = false;
    setIsLoadingSources(true);
    setError('');

    fetch(`${serverUrl}/api/knowledge-sources/options`, {
      method: 'GET',
      credentials: 'omit',
      headers: { Accept: 'application/json' }
    })
      .then(async (response) => {
        if (!response.ok) {
          throw new Error(`Knowledge sources could not be loaded (HTTP ${response.status}).`);
        }
        return response.json() as Promise<IKnowledgeSource[]>;
      })
      .then((items) => {
        if (cancelled) return;
        const enabledSources = Array.isArray(items) ? items : [];
        setSources(enabledSources);
        const portalWide = enabledSources.find((source) => source.isPortalWide);
        const initial = portalWide || enabledSources[0];
        setSelectedSourceId(initial ? String(initial.id) : '');
        if (!initial) {
          setError('No knowledge sources are currently available.');
        }
      })
      .catch((loadError: Error) => {
        if (!cancelled) {
          setError(loadError.message || 'Knowledge sources could not be loaded.');
        }
      })
      .then(() => {
        if (!cancelled) setIsLoadingSources(false);
      });

    return () => {
      cancelled = true;
    };
  }, [serverUrl]);

  const selectedSource = sources.find(
    (source) => String(source.id) === selectedSourceId
  );

  const submitQuestion = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    const message = question.trim();
    if (!message || !selectedSourceId || isSending) return;

    const userMessage: IChatMessage = {
      id: `user-${Date.now()}`,
      role: 'user',
      content: message
    };
    setMessages((current) => current.concat(userMessage));
    setQuestion('');
    setError('');
    setIsSending(true);

    try {
      const response = await fetch(`${serverUrl}/api/chat`, {
        method: 'POST',
        credentials: 'omit',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message,
          sourceId: Number(selectedSourceId),
          sessionId: sessionId.current,
          ...(loginName ? { username: loginName } : {})
        })
      });

      const result = await response.json();
      if (!response.ok) {
        throw new Error(result && result.message
          ? result.message
          : `The assistant could not answer (HTTP ${response.status}).`);
      }

      setMessages((current) => current.concat({
        id: `assistant-${result.id || Date.now()}`,
        role: 'assistant',
        content: String(result.content || 'No response was returned.')
      }));
    } catch (sendError) {
      const messageText = sendError instanceof Error
        ? sendError.message
        : 'The assistant could not answer. Please try again.';
      setError(messageText);
    } finally {
      setIsSending(false);
    }
  };

  return (
    <section className={styles.card} aria-label="inSite Assistant">
      <header className={styles.header}>
        <div className={styles.brandMark} aria-hidden="true">
          <span className={styles.brandBubble}>••</span>
        </div>
        <div>
          <div className={styles.eyebrow}>TECHNICAL ASSURANCE</div>
          <h2 className={styles.title}>inSite Assistant</h2>
        </div>
      </header>

      <div className={styles.body}>
        <p className={styles.description}>
          Ask a question using approved content from your SharePoint portal.
          {displayName ? ` Welcome, ${displayName}.` : ''}
        </p>

        <div className={styles.field}>
          <label htmlFor="inSite-source">Knowledge source</label>
          <select
            id="inSite-source"
            value={selectedSourceId}
            onChange={(event) => setSelectedSourceId(event.target.value)}
            disabled={isLoadingSources || isSending || sources.length === 0}
          >
            {isLoadingSources && <option value="">Loading sources…</option>}
            {!isLoadingSources && sources.length === 0 && (
              <option value="">No sources available</option>
            )}
            {sources.map((source) => (
              <option key={source.id} value={String(source.id)}>
                {source.name}
              </option>
            ))}
          </select>
          {selectedSource && (
            <span className={styles.helper}>
              {selectedSource.description || (
                selectedSource.isPortalWide
                  ? 'Searches all enabled portal knowledge sources.'
                  : `Searches only ${selectedSource.name}.`
              )}
            </span>
          )}
        </div>

        {messages.length > 0 && (
          <div className={styles.transcript} aria-live="polite">
            {messages.map((message) => (
              <div
                key={message.id}
                className={message.role === 'user' ? styles.userMessage : styles.assistantMessage}
              >
                <strong>{message.role === 'user' ? 'You' : 'inSite Assistant'}</strong>
                <span>{message.content}</span>
              </div>
            ))}
          </div>
        )}

        <form onSubmit={submitQuestion} className={styles.questionForm}>
          <label htmlFor="inSite-question">Your question</label>
          <div className={styles.questionRow}>
            <textarea
              id="inSite-question"
              value={question}
              onChange={(event) => setQuestion(event.target.value)}
              placeholder="What would you like to know?"
              rows={2}
              disabled={isSending}
            />
            <button
              type="submit"
              disabled={!question.trim() || !selectedSourceId || isSending}
            >
              {isSending ? 'Asking…' : 'Ask'}
              <span aria-hidden="true">→</span>
            </button>
          </div>
        </form>

        {error && (
          <div className={styles.error} role="alert">
            {error}
          </div>
        )}
      </div>
    </section>
  );
};

export default SourceAwareChat;