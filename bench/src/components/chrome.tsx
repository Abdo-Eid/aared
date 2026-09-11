import type { ButtonHTMLAttributes, ReactNode } from "react";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(...inputs));
}

interface IconButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  label: string;
  pressed?: boolean;
  size?: "header" | "footer";
}

/** Ghost icon button shared by the header and the footer. */
export function IconButton({
  label,
  pressed,
  size = "header",
  className,
  children,
  ...rest
}: IconButtonProps) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      aria-pressed={pressed}
      {...rest}
      className={cn(
        "focus-ring inline-flex shrink-0 cursor-pointer items-center justify-center transition-colors",
        "text-muted-foreground hover:bg-muted hover:text-foreground",
        "disabled:pointer-events-none disabled:opacity-40",
        size === "header"
          ? "size-8 rounded-[0.5625rem]"
          : "size-[1.875rem] rounded-lg",
        pressed === true &&
          "bg-accent-blue/10 text-accent-blue hover:bg-accent-blue/15",
        className,
      )}
    >
      {children}
    </button>
  );
}

interface IconLinkProps {
  label: string;
  href: string;
  download?: string;
  children: ReactNode;
}

/** An action that is a link — the download. */
export function IconLink({ label, href, download, children }: IconLinkProps) {
  return (
    <a
      href={href}
      aria-label={label}
      title={label}
      download={download}
      className="focus-ring inline-flex size-8 shrink-0 cursor-pointer items-center justify-center rounded-[0.5625rem] text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
    >
      {children}
    </a>
  );
}

/**
 * Solid square badge naming the format. Fills are literal so they clear 4.5:1
 * against white at badge size.
 */
export function FileBadge({
  className,
  label,
}: {
  className: string;
  label: string;
}) {
  return (
    <span
      className={cn(
        "flex size-9 shrink-0 items-center justify-center rounded-lg font-mono text-[0.6875rem] font-medium uppercase leading-none tracking-tight text-white",
        className,
      )}
    >
      {label}
    </span>
  );
}

/** Stage-sized status message: opening, empty, or unable to open. */
export function StageMessage({
  icon,
  children,
  className,
}: {
  icon?: ReactNode;
  children: ReactNode;
  className?: string;
}) {
  return (
    <div
      className={cn(
        "flex min-w-0 flex-1 flex-col items-center justify-center gap-2 bg-card px-6 text-center text-sm text-muted-foreground",
        className,
      )}
      role="status"
    >
      {icon}
      <p className="max-w-xs">{children}</p>
    </div>
  );
}

/** Hairline divider for the footer bar. */
export function Separator() {
  return (
    <span aria-hidden="true" className="mx-[0.3125rem] h-4 w-px bg-border" />
  );
}
