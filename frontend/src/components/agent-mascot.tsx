import Image from "next/image";

import { cn } from "@/lib/utils";

type AgentMascotProps = {
  className?: string;
  priority?: boolean;
};

export function AgentMascot({ className, priority = false }: AgentMascotProps) {
  return (
    <span className={cn("relative block shrink-0", className)} aria-hidden="true">
      <Image
        src="/images/hr-owl-mascot.png"
        alt=""
        fill
        priority={priority}
        sizes="(max-width: 640px) 20px, 32px"
        className="object-contain object-center"
      />
    </span>
  );
}
