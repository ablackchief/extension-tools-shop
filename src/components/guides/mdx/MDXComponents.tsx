import type { MDXComponents } from 'mdx/types';
import { Callout } from './Callout';
import { Stat } from './Stat';
import { ExampleBox } from './ExampleBox';
import { ToolCTA } from './ToolCTA';

export const mdxComponents: MDXComponents = {
  Callout,
  Stat,
  ExampleBox,
  ToolCTA,
  // Override default elements for consistent styling
  h2: ({ children, ...props }) => {
    const id = typeof children === 'string'
      ? children.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '')
      : undefined;
    return (
      <h2
        id={id}
        className="text-2xl font-bold text-slate-900 mt-12 mb-4 scroll-mt-24"
        {...props}
      >
        {children}
      </h2>
    );
  },
  h3: ({ children, ...props }) => (
    <h3
      className="text-xl font-bold text-slate-900 mt-8 mb-3"
      {...props}
    >
      {children}
    </h3>
  ),
  h4: ({ children, ...props }) => (
    <h4
      className="text-lg font-semibold text-slate-900 mt-6 mb-2"
      {...props}
    >
      {children}
    </h4>
  ),
  p: ({ children, ...props }) => (
    <p className="text-slate-600 leading-relaxed mb-4" {...props}>
      {children}
    </p>
  ),
  ul: ({ children, ...props }) => (
    <ul className="list-disc list-inside text-slate-600 mb-4 space-y-1 ml-4" {...props}>
      {children}
    </ul>
  ),
  ol: ({ children, ...props }) => (
    <ol className="list-decimal list-inside text-slate-600 mb-4 space-y-1 ml-4" {...props}>
      {children}
    </ol>
  ),
  li: ({ children, ...props }) => (
    <li className="text-slate-600" {...props}>
      {children}
    </li>
  ),
  strong: ({ children, ...props }) => (
    <strong className="font-semibold text-slate-900" {...props}>
      {children}
    </strong>
  ),
  blockquote: ({ children, ...props }) => (
    <blockquote
      className="border-l-4 border-[var(--primary-500)] bg-slate-50 py-2 px-6 my-6 rounded-r-lg italic text-slate-700"
      {...props}
    >
      {children}
    </blockquote>
  ),
  a: ({ children, href, ...props }) => (
    <a
      href={href}
      className="text-[var(--primary-600)] hover:underline"
      {...props}
    >
      {children}
    </a>
  ),
  hr: () => (
    <hr className="my-8 border-slate-200" />
  ),
};
