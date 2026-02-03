import fs from 'fs';
import path from 'path';
import matter from 'gray-matter';
import readingTime from 'reading-time';

const guidesDirectory = path.join(process.cwd(), 'content/guides');

export interface ToolCTA {
  name: string;
  description: string;
  bundle: string;
  link: string;
}

export interface FAQ {
  question: string;
  answer: string;
}

export interface GuidePost {
  slug: string;
  title: string;
  description: string;
  author: string;
  authorCredential: string;
  authorBio: string;
  date: string;
  updated: string;
  category: string;
  tags: string[];
  readingTime: number;
  toolCTA: ToolCTA;
  relatedPosts: string[];
  featuredImage: string;
  featuredImageAlt: string;
  faqs: FAQ[];
  published: boolean;
  content: string;
}

export function getAllGuides(): GuidePost[] {
  // Ensure directory exists
  if (!fs.existsSync(guidesDirectory)) {
    return [];
  }

  const fileNames = fs.readdirSync(guidesDirectory);
  const allGuides = fileNames
    .filter((fileName) => fileName.endsWith('.mdx'))
    .map((fileName) => {
      const slug = fileName.replace(/\.mdx$/, '');
      const fullPath = path.join(guidesDirectory, fileName);
      const fileContents = fs.readFileSync(fullPath, 'utf8');
      const { data, content } = matter(fileContents);
      const stats = readingTime(content);

      return {
        slug,
        title: data.title || '',
        description: data.description || '',
        author: data.author || '',
        authorCredential: data.authorCredential || '',
        authorBio: data.authorBio || '',
        date: data.date || '',
        updated: data.updated || data.date || '',
        category: data.category || '',
        tags: data.tags || [],
        readingTime: data.readingTime || Math.ceil(stats.minutes),
        toolCTA: data.toolCTA || {
          name: '',
          description: '',
          bundle: '',
          link: '/tools/',
        },
        relatedPosts: data.relatedPosts || [],
        featuredImage: data.featuredImage || '',
        featuredImageAlt: data.featuredImageAlt || '',
        faqs: data.faqs || [],
        published: data.published !== false,
        content,
      } as GuidePost;
    })
    .filter((guide) => guide.published)
    .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

  return allGuides;
}

export function getGuideBySlug(slug: string): GuidePost | null {
  try {
    const fullPath = path.join(guidesDirectory, `${slug}.mdx`);

    if (!fs.existsSync(fullPath)) {
      return null;
    }

    const fileContents = fs.readFileSync(fullPath, 'utf8');
    const { data, content } = matter(fileContents);
    const stats = readingTime(content);

    return {
      slug,
      title: data.title || '',
      description: data.description || '',
      author: data.author || '',
      authorCredential: data.authorCredential || '',
      authorBio: data.authorBio || '',
      date: data.date || '',
      updated: data.updated || data.date || '',
      category: data.category || '',
      tags: data.tags || [],
      readingTime: data.readingTime || Math.ceil(stats.minutes),
      toolCTA: data.toolCTA || {
        name: '',
        description: '',
        bundle: '',
        link: '/tools/',
      },
      relatedPosts: data.relatedPosts || [],
      featuredImage: data.featuredImage || '',
      featuredImageAlt: data.featuredImageAlt || '',
      faqs: data.faqs || [],
      published: data.published !== false,
      content,
    } as GuidePost;
  } catch {
    return null;
  }
}

export function getRelatedGuides(slugs: string[]): GuidePost[] {
  return slugs
    .map((slug) => getGuideBySlug(slug))
    .filter((guide): guide is GuidePost => guide !== null);
}

export function getGuideCategories(): string[] {
  const guides = getAllGuides();
  const categories = new Set(guides.map((guide) => guide.category));
  return Array.from(categories).sort();
}

export function getGuidesByCategory(category: string): GuidePost[] {
  const guides = getAllGuides();
  if (category === 'all') return guides;
  return guides.filter((guide) => guide.category === category);
}
